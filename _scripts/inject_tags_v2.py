# -*- coding: utf-8 -*-
"""form/templates/*.hwpx 에 {{태그}} 를 심는다. (신 서식 기준)

한글 COM 을 쓰지 않고 hwpx(zip+XML)를 직접 고친다. 이유가 셋 있다.
  · 한글 찾아바꾸기는 '몇 번째 것' 을 지정할 수 없다. 순회점검표의 1주~5주 점검일은
    똑같은 '0000.00.00.' 이 다섯 번 나오므로 n번째를 각각 다른 태그로 바꿔야 한다.
  · 값 칸이 비어 있는 표(계약이행 서약서 등)는 찾아바꾸기로 채울 수 없다.
    XML 에서는 라벨 칸의 cellAddr 을 읽어 그 오른쪽/아래 칸을 정확히 집을 수 있다.
  · 한글을 띄우지 않으니 빠르고, 결과가 늘 같다.

표의 행·칸 구조는 절대 건드리지 않는다. 글자만 바꾸거나 빈 칸에 써 넣는다.

사용법:
    python _scripts/inject_tags_v2.py --dry-run   # 무엇을 바꿀지만 출력
    python _scripts/inject_tags_v2.py             # 실제로 심고 결과 확인
"""
import os
import re
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TDIR = os.path.join(ROOT, "form", "templates")

# ---------------------------------------------------------------- 작업 목록
#
#   ("t",     옛 글자, 새 글자)            글자가 든 run 을 찾아 바꾼다 (모두)
#   ("seq",   옛 글자, [새1, 새2, …])      똑같은 글자의 1·2·3…번째를 각각 다르게
#   ("right", 라벨 칸 글자, 태그)          그 칸 오른쪽 칸에 써 넣는다
#   ("below", 라벨 칸 글자, 태그)          그 칸 아래 칸에 써 넣는다
#   ("para",    머리글자, 새 문단 글자)     그 글자로 시작하는 첫 문단을 통째로 바꾼다
#   ("paraseq", 머리글자, [새1, 새2, …])    그 글자로 시작하는 문단들을 차례로 바꾼다
#   ("dup",     문단 글자, [새1, 새2, …])   그 문단을 복제해 바로 아래에 줄을 만든다
#   ("indent",  머리글자, 왼쪽여백)         그 문단들에 내어쓰기를 준다
#
# 라벨은 칸 글자 전체와 똑같아야 집힌다. 본문에 같은 낱말이 나와도 걸리지 않는다.
#
# "t" 는 한 <hp:t> 안에서만 찾으므로 글자가 여러 run 으로 쪼개져 있으면 못 잡는다.
# 매뉴얼 문서는 한 줄이 서너 개의 run 으로 쪼개져 있는 곳이 많다. 그래서 파란 예시
# 글을 갈아 끼울 때는 문단 단위로 보는 "para" 계열을 쓴다. 머리글자는 앞 공백을 뗀
# 뒤에 견주므로 ' - 작업절차:' 든 '- 작업절차:' 든 똑같이 걸린다.

OPS = {
    # ── 09 작업장 순회점검표 ────────────────────────────────────────────
    # 신 양식은 '한 달에 한 장' 이다. 제목에 년·월이 있고 표 안에 1주~5주 행이 있다.
    # 그 달의 점검일을 각 주차 칸에 따로 채워야 하므로 다섯 개를 각각 다른 태그로 바꾼다.
    "09_순회점검표.hwpx": [
        ("t", "작업장 순회점검표 [     년     월]",
              "작업장 순회점검표 [{{점검년}}년 {{점검월}}월]"),
        ("seq", "0000.00.00.", ["{{점검일1}}", "{{점검일2}}", "{{점검일3}}",
                                "{{점검일4}}", "{{점검일5}}"]),
    ],

    # ── 01 안전보건관리계획서(상세) ─────────────────────────────────────
    # ① 끝에 공백이 붙은 태그와, 다른 서식과 이름이 어긋난 태그를 맞춘다.
    # ② ③ 위험성평가 칸의 '○○ ○○ 설치를 위한 …' 같은 매뉴얼 예시 글을
    #    계약유형 사전에서 오는 태그로 바꾼다. 계약명을 넣으면 그 계약에 맞는
    #    작업내용·위험요인·안전대책이 채워진다.
    "01_계획서_상세.hwpx": [
        ("t", "{{업체대표자연락처 }}", "{{업체대표자연락처}}"),
        ("t", "{{소방서}}", "{{관할소방서}}"),
        ("t", "{{병원}}", "{{인근병원}}"),

        # ③ 위험성평가 — 작업 중 유해·위험 사항
        # 순서가 중요하다. 긴 글부터 지워야 짧은 조각(' ② ' 같은 것)을 지울 때
        # 살아 있어야 할 곳(② 안전보건관리활동계획 머리글)까지 지우지 않는다.
        ("t", "작업명: ○○ ○○ 설치를 위한 비계설치", "작업명: {{작업내용}}"),
        ("t", "·해체 작업", ""),
        ("t", "  ① 추락 위험", "  ① {{위험요인}}"),
        ("t", "  ② 낙하물 위험", ""),
        ("t", " 2. 작업명: 용접 및 절단 작업", ""),
        ("t", "  ① 용접 시 화재위험", ""),
        ("t", "  ② 절단 시 부상위험", ""),
        ("t", " 3. 작업명: 페인트칠 작업", ""),
        ("t", "  - 유해물질 흡입 위험", ""),

        # ③ 위험성평가 — 안전보건관리 대책
        ("t", "작업명:  ○○ ○○ 설치를 위한  비계설치", "작업명: {{작업내용}}"),
        ("seq", " ① 개인보호구 지급 및 착용", [" ① {{안전대책}}", ""]),
        ("t", " ② 추락위험장소의 안전시설물 설치", ""),
        ("t", " ③ 무리한 해체 작업 금지", ""),
        ("t", " ④ 작업자 숙련 정도에 따라 적정배치", ""),
        ("t", "2인 1조 작업계획 확인", ""),
        ("t", "작업명: 용접 및 절단 작업", ""),
        ("t", "소화기 및 용접방화포 등 방화용품 비치 ", ""),
        ("t", "3. 작업명: 페인트칠 작업", ""),
        ("t", "보호구(마스크, 장갑, 안전화) 착용", ""),
        # 위에서 글을 지우고 남은 번호 조각. ② ③ 은 위쪽 머리글에도 쓰이므로
        # **그것만 남은 문단**을 골라 지운다. 예전에는 글자가 똑같은 run 을 지웠는데,
        # 그러면 ' - 현장 종사자수: ○명' 처럼 ' - ' 가 따로 run 으로 떨어져 있는
        # 멀쩡한 줄에서 머리표까지 지워 버린다.
        ("clearpara", "⑤"),
        ("clearpara", "②"),
        ("clearpara", "③"),
        ("clearpara", "-"),

        # ② 사용 기계 기구 및 설비 ― 계약유형 사전에서 가져온다
        ("para", " - 용접기:", " - {{사용기계기구1}}"),
        ("para", " - 절단기:", " - {{사용기계기구2}}"),
        ("para", " - 비계:", " - {{사용기계기구3}}"),

        # ⑥ 위험물질 및 설비
        ("para", "외부 4면 시스템비계 설치", "{{위험물질설비1}}"),
        ("para", "용접기 및 절단기", "{{위험물질설비2}}"),
        ("para", "화학제품 등", "{{위험물질설비3}}"),

        # ① 인원 수는 엑셀에 없는 값이라 손으로 채우는 칸으로 남긴다.
        #    '(파란색 수정)' 은 매뉴얼이 작성자에게 남긴 말이라 지운다.
        ("para", " - 현장 종사자수:", " - 현장 종사자수: ○명"),

        # ③ 위험성평가 ― 한 줄에 한 항목씩.
        # 오른쪽 칸의 '작업명:' 문단에는 **한글의 자동 문단번호**가 붙어 있다.
        # 여기에 ' 1. ' 을 글자로 또 넣으면 '1.  1. 작업명:' 이 되어 번호가 두 번
        # 찍힌다. 눈으로는 왼쪽 칸과 같아 보이지만 글자는 다르다. 손대지 않는다.
        ("para", "① {{위험요인}}", " ① {{위험요인1}}"),
        ("para", "① {{안전대책}}", " ① {{안전대책1}}"),
        # ④ 는 쓰지 않는다. 학교 현장 항목(역할 분담·권장 시기·위험성 등급)을
        # 넣으면서 두 쪽을 넘기지 않으려고 ③ 까지로 줄였다. 사전에는 4개가
        # 그대로 있으므로, 나중에 자리가 생기면 ④ 를 도로 넣기만 하면 된다.
        # 원 번호 앞의 공백을 뗀다. 들여쓰기는 아래 ("indent", …, 2750, 1100) 이
        # 문단모양으로 준다. 공백으로 맞추면 판마다 자리가 달라진다.
        ("striplead", "① {{"),
        ("striplead", "② {{"),
        ("striplead", "③ {{"),
        ("dup", "① {{위험요인1}}", ["② {{위험요인2}}",
                                    "③ {{위험요인3}}"]),
        ("dup", "① {{안전대책1}}", ["② {{안전대책2}}",
                                    "③ {{안전대책3}}"]),

        # 내어쓰기 ― 줄이 감길 때 다음 줄이 항목 글자 시작 위치에 맞춰 들어간다.
        # 이것이 없으면 감긴 줄이 칸 왼쪽 끝으로 붙어 번호와 글이 뒤엉킨다.
        #
        # ①②③④ 항목에만 준다. '작업명:' 줄과 ② 기계기구 줄에도 걸어 보았더니,
        # 칸이 좁아 글자가 들어갈 자리를 내어쓰기가 먹어 버려 오히려 줄이 늘어났다.
        # 그 두 곳은 ① 인력계획 칸과 같이 내어쓰기 없이 둔다.
        # 공문서 방식 ― 원 번호는 스페이스 두 칸 자리(1100)에서 시작하고, 감긴 줄은
        # 번호 뒤 글자 자리(2750 = 두 칸 + 전각 ① + 한 칸)에 맞춰 들어간다.
        ("indent", "① {{", 2750, 1100),
        ("indent", "② {{", 2750, 1100),
        ("indent", "③ {{", 2750, 1100),

        # ── 학교 현장 항목 ────────────────────────────────────────────────
        # 두 쪽을 넘기면 안 되므로 넣은 줄과 뺀 줄을 같은 수로 맞췄다.
        #   넣음 +8 : 역할 분담 6줄, 권장 시기 1줄, 위험성 등급 1줄
        #   뺌  -8 : ① 연락처 중복 4줄, ③ ④번 항목 2줄, ⑤ 중복 1줄, 안내문 1줄
        #
        # □ 계약현황 ― 언제 할 일인지와 얼마나 위험한지를 맨 앞에서 보여 준다.
        ("dup", "◦계약기간 : {{계약기간}}", ["◦권장 시기: {{작업권장시기}}",
                                             "◦위험성 등급: {{위험성등급}}"]),

        # ① '역할 분담' 은 제목만 있고 정작 내용이 없던 자리다. 학교가 확인할 것과
        #    업체가 조치할 것을 갈라 적는다. 계획서에서 가장 필요한 대목이다.
        # '학교' '업체' 를 줄마다 되풀이하지 않고 두 줄로 묶는다. 항목도 둘씩만 쓴다
        # (사전에는 셋이 그대로 있다). 오른쪽 칸과 줄 수를 맞추기 위해서다.
        ("dup", "◦이행계획 추진을 위한 구성원의 역할 분담",
                [" - 학교: {{학교확인사항1}}, {{학교확인사항2}}",
                 " - 업체: {{업체조치사항1}}, {{업체조치사항2}}"]),
        # 예전에 한 줄씩 넣어 둔 여섯 줄을 걷어낸다. 이미 없으면 그냥 지나간다.
        ("clearpara", "- 학교: {{학교확인사항1}}"),
        ("clearpara", "- 학교: {{학교확인사항2}}"),
        ("clearpara", "- 학교: {{학교확인사항3}}"),
        ("clearpara", "- 업체: {{업체조치사항1}}"),
        ("clearpara", "- 업체: {{업체조치사항2}}"),
        ("clearpara", "- 업체: {{업체조치사항3}}"),

        # ① 칸의 왼쪽·오른쪽 덩어리를 맞바꾼다. 왼쪽이 오른쪽보다 훨씬 길어
        # 칸 하나만 늘어져 보였다. '역할 분담' 을 오른쪽으로, '활동계획' 을
        # 왼쪽으로 보내면 양쪽 줄 수가 비슷해진다.
        # 차례가 중요하다. 역할 분담을 먼저 오른쪽 끝으로 보낸 뒤 활동계획을
        # 데려와야 오른쪽 칸이 한 번도 비지 않는다.
        ("movepara", ["◦이행계획 추진을 위한 구성원의 역할 분담",
                      "- 학교:", "- 업체:"], "·그 밖의 인원:"),
        ("movepara", ["◦활동계획: 일 1회 이상 안전점검", "- 현장 종사자수:",
                      "·세부작업 인원:", "·그 밖의 인원:"], "- 공정별 2인 1조"),

        # 태그 자리는 파란 글씨로 ― 그대로 두지 말고 확인하라는 뜻이다.
        ("tagcolor", "#0000FF"),

        # ⑤ 신호·연락체계에 그대로 또 있는 줄. 같은 {{안전담당자}} 가 세 번 찍혔다.
        ("clearpara", "◦현장책임자: {{안전담당자}}"),
        ("clearpara", "◦안전담당: {{안전담당자}}"),
        ("clearpara", "- 연락처: {{안전담당자연락처}}"),
        # ⑤ 2번(현장책임자)과 3번(안전담당)이 글자까지 같다. 하나만 남기고 번호를 당긴다.
        ("clearpara", "3. 안전담당: {{안전담당자}} ☎ {{안전담당자연락처}}"),
        ("para", "4. 행정실 업무담당자",
                 "3. 행정실 업무담당자: {{행정실담당자}} ☎ {{행정실전화번호}}"),
        # ③ ④번 항목과 ④ 머리의 매뉴얼 안내문
        ("clearpara", "④ {{위험요인4}}"),
        ("clearpara", "④ {{안전대책4}}"),
        ("clearpara", "‘별도의 교육자료 첨부가능’"),
    ],

    # ── 01 안전보건관리계획서(공사) ─────────────────────────────────────
    # 파란 글씨는 매뉴얼의 예시(논슬립 교체)가 그대로 박힌 것이다. 논슬립 공사가
    # 아닌 계약에도 그 글이 나오므로 계약유형 사전에서 오는 태그로 바꾼다.
    # 한 쪽짜리 서식이라 줄 수는 늘리지 않는다 ― 위험요인 2줄, 관리계획 2줄 그대로.
    "01_계획서_공사.hwpx": [
        ("para", " - 작업절차:", " - 작업절차: {{작업절차}}"),
        ("para", " - 안전점검계획:", " - 안전점검계획: {{안전점검계획}}"),
        ("para", "◦ 보호구 및 방호장치계획:",
                 "◦ 보호구 및 방호장치계획: {{보호구방호장치}}"),
        # '·' 로 시작하는 문단은 유해위험요인 2줄 · 관리계획 2줄, 이 차례로 나온다
        ("paraseq", "·", [" ·{{위험요인1}}", " ·{{위험요인2}}",
                          " ·{{안전대책1}}", " ·{{안전대책2}}"]),
        ("para", " - 작업전 기계", " - {{안전보건교육}}"),
        ("indent", "·{{", 1100),
    ],

    # ── 01 안전보건관리계획서(용역) ─────────────────────────────────────
    # 공사용과 짜임이 같다. 예시가 음식물쓰레기 수거인 것만 다르다.
    "01_계획서_용역.hwpx": [
        ("para", " - 작업절차:", " - 작업절차: {{작업절차}}"),
        ("para", " - 안전점검계획:", " - 안전점검계획: {{안전점검계획}}"),
        ("para", "◦ 보호구 및 방호장치계획:",
                 "◦ 보호구 및 방호장치계획: {{보호구방호장치}}"),
        ("paraseq", "·", [" ·{{위험요인1}}", " ·{{위험요인2}}",
                          " ·{{안전대책1}}", " ·{{안전대책2}}"]),
        ("para", " - 운반물 취급시", " - {{안전보건교육}}"),
        ("indent", "·{{", 1100),
    ],

    # ── 03 안전보건협의체 회의록 ────────────────────────────────────────
    # 협의사항 칸의 예시(분리수거·사다리)를 계약유형 사전 문구로 바꾼다.
    # 2번·5번 항목의 {{학교명}} {{행정실전화번호}} {{업체담당자}} 는 손대지 않는다.
    "08_협의체회의록.hwpx": [
        # 2번 연락처는 학교·업체 두 줄이 한 문단 안에 줄바꿈으로 들어 있었다.
        # 그러면 업체 줄만 내어쓰기만큼 더 들어가 보인다. 문단을 갈라 놓는다.
        ("splitbr", "∙ {{학교명}} 교육행정실"),
        # 가운뎃점 앞의 공백 넉 칸을 뗀다. 들여쓰기는 아래 ("indent", "∙ ", …) 가
        # 문단모양으로 준다.
        ("striplead", "∙ "),

        ("para", "202○.", "   {{협의체회의일}}  {{계약명}}"),
        ("para", "비상 시 비상계단", "   {{대피방법}}"),
        ("para", "예) 분리수거 중", "  예) {{위험요인1}}"),
        # 4번의 '∙' 줄(위험요인2·안전대책1)은 이미 심겨 있다. 예전에는 여기서
        # ("paraseq", "∙ ", […]) 로 넣었는데, 그 조작을 남겨 두면 다시 돌릴 때
        # **2번 항목의 연락처 줄**부터 '∙' 로 시작한다는 이유로 덮어써 버린다.
        # (학교 행정실·업체 담당자 전화번호가 통째로 사라진다.) 그래서 지웠다.
        # 위험요인③·안전대책② 는 아래에서 뺀다. 6번(의결사항)을 넣을 자리다.

        # ── 학교 현장 항목 ────────────────────────────────────────────────
        # 넣은 줄과 뺀 줄을 같은 수로 맞췄다(+4 / -4). 한 쪽을 넘기지 않는다.
        #
        # 2번 ― '유선전화 사용하여 상호 간 연락' 은 바로 위 두 줄에 이미 번호가
        #        적혀 있어 뜻이 없는 상투구다. 그 자리를 상시 안건으로 바꾼다.
        # 머리글자를 짧게 잡는다. 이 줄은 '상호 간' 앞뒤가 다른 run 으로 쪼개져
        # 있어서 문장 전체로는 걸리지 않는다.
        ("para", "∙ 유선전화", "∙ {{상시안건1}}"),
        # 4번 ― 매뉴얼이 작성자에게 남긴 안내문. 출력물에 남을 글이 아니다.
        ("clearpara", "∙ {{위험요인3}}"),
        ("clearpara", "∙ {{안전대책2}}"),
        ("clearpara", "(당일 작업에 따른 유해ㆍ위험요인에 따라 안전대책 작성)"),
        ("clearpara", "(위험성평가 참고하여 작성)"),
        # 6번 ― 지금 회의록에는 협의만 있고 의결이 없다. 합의한 결과를 적는 자리를
        #        서명표 바로 위에 만든다.
        ("dup", "∙ 업체 담당자 {{업체담당자}}, {{업체담당자연락처}} 필요 시 유선 연락",
                [" 6. 협의 결과(의결사항)",
                 "∙ {{협의체합의문1}}",
                 "∙ {{협의체합의문2}}",
                 "∙ {{협의체합의문3}}"]),
        ("indent", "6. 협의 결과", 0),

        # 가운뎃점 항목이 넉 칸이나 들어가 있었다(글자 앞 공백 4 + 내어쓰기 4400).
        # 공문서 방식대로 숫자 항목보다 두 칸만 들어가게 한다.
        ("indent", "∙ ", 2200, 1100),
        # 태그 자리는 파란 글씨로 ― 그대로 두지 말고 확인하라는 뜻이다.
        ("tagcolor", "#0000FF"),

        ("indent", "{{협의체회의일}}", 1650),
        # 예전에 있던 ("indent", "∙ {{", 4400) 은 지웠다. 위의 ("indent", "∙ ", …)
        # 보다 뒤에 있어서 새 값을 도로 4400 으로 덮어썼다.
        # ("indent", "{{대피방법}}") 과 ("indent", "예) {{위험요인1}}") 도 지웠다.
        # 그 줄들은 이제 '∙ ' 로 시작해 위의 한 줄이 함께 맡는다.
    ],

    # ── 11 아차사고 사례 제안서 ─────────────────────────────────────────
    "11_아차사고제안서.hwpx": [
        ("t", "학 교 명 :    ", "학 교 명 : {{학교명}}"),
        ("t", "일 시 : ", "일 시 : {{작성일}}"),
    ],

    # ── 12 관리감독자 업무수행 평가표 ───────────────────────────────────
    "12_관리감독자평가표.hwpx": [
        ("t", "교(기관)명:", "교(기관)명: {{학교명}}"),
        ("t", "직:             성명:             (인)",
              "직: {{관리자직위}}    성명: {{관리자성명}}    (인)"),
        ("t", "평가일자:      .     .       . ", "평가일자: {{작성일}}"),
    ],

    # ── 13 중대재해발생보고서 ───────────────────────────────────────────
    # 사고 내용은 사람이 쓰는 것이라 비워 둔다. 기관·작성자만 채운다.
    "13_중대재해발생보고서.hwpx": [
        # 태그 글자가 원래 빈칸보다 길면 줄이 넘어가 표가 다음 쪽으로 밀린다.
        # 값이 들어가면 다시 짧아지지만, 템플릿을 열어 봤을 때도 한 쪽이어야 한다.
        # 사용자가 한글에서 두 줄을 한 줄로 합쳐 두었다. 새로 자른 서식에서도
        # 같은 모양이 나오도록 바꿀 글자를 그 모양에 맞춘다.
        ("t", "2000.    .    .    작 성 자: 직)                 성명)   ",
              "작성자: {{담당자직위}} {{담당자성명}},"),
        ("t", "                               확 인 자: 직)                 성명)",
              " 확인자: {{부서장직위}} {{부서장성명}}"),
        ("below", "사업장명", "{{학교명}}"),
        ("below", "대표자", "{{기관대표자}}"),
        # 소재지 칸은 좁아서 주소를 넣으면 줄이 넘어가고 표가 다음 쪽으로 밀린다.
        # 사고 보고서는 어차피 손으로 채우는 칸이 많으므로 비워 둔다.
    ],

    # ── 14 중대재해사고보고서 ───────────────────────────────────────────
    "14_중대재해사고보고서.hwpx": [
        # 태그 글자가 원래 빈칸보다 길면 줄이 넘어가 표가 다음 쪽으로 밀린다.
        # 값이 들어가면 다시 짧아지지만, 템플릿을 열어 봤을 때도 한 쪽이어야 한다.
        # 사용자가 한글에서 두 줄을 한 줄로 합쳐 두었다. 새로 자른 서식에서도
        # 같은 모양이 나오도록 바꿀 글자를 그 모양에 맞춘다.
        ("t", "2000.    .    .    작 성 자: 직)                 성명)   ",
              "작성자: {{담당자직위}} {{담당자성명}},"),
        ("t", "                               확 인 자: 직)                 성명)",
              " 확인자: {{부서장직위}} {{부서장성명}}"),
        ("right", "사고현장", "{{학교명}}"),
        ("right", "계약명", "{{계약명}}"),
        ("right", "계약자", "{{기관명}}"),
        ("right", "계약기간", "{{계약기간}}"),
        ("right", "계약상대자", "{{업체명}}"),
        ("right", "계약금액", "{{계약금액}}"),
    ],

    # ── 15 계약이행 통합 서약서 ─────────────────────────────────────────
    "15_계약이행서약서.hwpx": [
        ("right", "계약명", "{{계약명}}"),
        ("right", "발주기관", "{{기관명}}"),
        ("right", "업체명", "{{업체명}}"),
        ("right", "사업자등록번호", "{{사업자등록번호}}"),
        ("right", "연락처", "{{업체전화}}"),
        ("right", "주소", "{{업체주소}}"),
        # 대표자 칸은 '(인)' 이 이미 들어 있어 비어 있지 않다. 글자를 바꿔 넣는다.
        ("t", "                     (인)", "{{업체대표자}}            (인)"),
        ("t", "202  .   .   .", "{{작성일}}"),
        ("t", "계약상대자:                대표            (인)",
              "계약상대자: {{업체명}}      대표 {{업체대표자}}      (인)"),
        ("t", "○○○○장 귀하", "{{기관명}} {{기관대표자직위}} 귀하"),
    ],
}


# ------------------------------------------------------------------ XML 도구

def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def read_section(path):
    with zipfile.ZipFile(path) as z:
        entries = {n: z.read(n) for n in z.namelist()}
        info = {i.filename: i for i in z.infolist()}
    return entries, info


def write_section(path, entries, info):
    """옆에 새 파일을 만든 뒤 바꿔치기한다.

    서식 파일에 바로 쓰다가 중간에 멈추면 파일이 반토막 난다. 되돌릴 원본이
    없으므로(사용자가 손수 만든 서식이다) 반드시 다 쓴 뒤에 바꾼다.
    """
    tmp = path + ".new"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            # mimetype 은 압축하지 않아야 한다(OCF 규약)
            ct = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
            zi = zipfile.ZipInfo(name, date_time=info[name].date_time)
            zi.compress_type = ct
            zi.external_attr = info[name].external_attr
            z.writestr(zi, data)
    os.replace(tmp, path)


T_RE = re.compile(r"<hp:t>(.*?)</hp:t>", re.S)


def replace_text(xml, old, new):
    """<hp:t> 안의 글자를 모두 바꾼다. (바뀐 수, 새 xml)"""
    old_x, new_x = xml_escape(old), xml_escape(new)
    hits = [0]

    def one(m):
        body = m.group(1)
        if old_x not in body:
            return m.group(0)
        hits[0] += body.count(old_x)
        return "<hp:t>" + body.replace(old_x, new_x) + "</hp:t>"

    out = T_RE.sub(one, xml)          # 먼저 돌린 뒤에 센 값을 읽어야 한다
    return hits[0], out


def replace_whole_run(xml, old, new):
    """글자가 old 와 **똑같은** run 만 바꾼다. (바뀐 수, 새 xml)

    ' ② ' 같은 짧은 조각을 지울 때 쓴다. 부분 일치로 지우면
    ' ② 안전·보건관리활동계획 ' 같은 머리글의 번호까지 날아간다.
    """
    old_x, new_x = xml_escape(old), xml_escape(new)
    hits = [0]

    def one(m):
        if m.group(1) != old_x:
            return m.group(0)
        hits[0] += 1
        return "<hp:t>" + new_x + "</hp:t>"

    out = T_RE.sub(one, xml)
    return hits[0], out


def replace_sequence(xml, old, news):
    """똑같은 글자가 여러 번 나올 때 1번째·2번째… 를 각각 다른 글자로 한 번에 바꾼다.

    한 번에 처리해야 한다. 하나씩 나눠 바꾸면 앞의 것이 사라져 순서가 밀린다.
    """
    old_x = xml_escape(old)
    news_x = [xml_escape(s) for s in news]
    seen = [0]
    hits = [0]

    def one(m):
        body = m.group(1)
        if old_x not in body:
            return m.group(0)
        out, pos = [], 0
        while True:
            i = body.find(old_x, pos)
            if i < 0:
                out.append(body[pos:])
                break
            out.append(body[pos:i])
            if seen[0] < len(news_x):
                out.append(news_x[seen[0]])
                hits[0] += 1
            else:
                out.append(old_x)
            seen[0] += 1
            pos = i + len(old_x)
        return "<hp:t>" + "".join(out) + "</hp:t>"

    out = T_RE.sub(one, xml)
    return hits[0], out


# ------------------------------------------------------- 문단 단위로 고치기
#
# 매뉴얼 문서는 한 줄이 서너 개의 run 으로 쪼개져 있다(라벨은 검은색, 값은
# 파란색, 가운뎃점만 또 다른 글꼴 …). 그래서 "글자 찾아 바꾸기" 로는 한 줄을
# 통째로 갈아 끼울 수 없다. 아래 함수들은 <hp:p> 를 단위로 본다.

P_RE = re.compile(r"<hp:p\b[^>]*>.*?</hp:p>", re.S)


def para_text(p):
    """문단 하나의 글자. run 이 몇 개로 쪼개져 있든 이어 붙인다."""
    return "".join(T_RE.findall(p))


def rewrite_para(p, new_x):
    """문단의 글자를 new_x 로 바꾼다. run 서식은 최대한 살린다.

    앞부분이 지금 글자와 같으면 그 부분의 run 은 손대지 않고, 달라지는 곳부터를
    그 자리의 run 하나에 몰아 넣는다. 뒤따르던 run 은 글자만 비운다(지우지 않는다
    ― run 안에 줄바꿈이나 개체가 들어 있을 수 있다).

    그래서 '- 작업절차: ' 같은 검은 라벨은 검은색으로 남고, 값이 들어가는 자리는
    원래 파란색이던 run 을 그대로 물려받아 파란색으로 나온다.
    """
    old_x = para_text(p)
    if old_x == new_x:
        return p, False

    ts = list(T_RE.finditer(p))
    if not ts:
        return p, False

    c = 0
    while c < min(len(old_x), len(new_x)) and old_x[c] == new_x[c]:
        c += 1

    bodies, pos, carrier = [], 0, None
    for i, m in enumerate(ts):
        body = m.group(1)
        start, end = pos, pos + len(body)
        pos = end
        if carrier is not None:
            bodies.append("")                         # 남은 run 은 글자만 비운다
        elif end <= c and end < len(old_x):
            bodies.append(body)                       # 앞부분은 그대로
        else:
            # 여기서부터 글자가 달라진다. 마지막 run 은 새 글자가 옛 글자보다
            # 길기만 할 때(옛 글자가 통째로 앞토막일 때)도 반드시 여기로 온다.
            carrier = i
            bodies.append(body[:max(0, c - start)] + new_x[c:])

    out, last = [], 0
    for m, body in zip(ts, bodies):
        out.append(p[last:m.start()])
        out.append("<hp:t>" + body + "</hp:t>")
        last = m.end()
    out.append(p[last:])
    return "".join(out), True


def para_spans(xml):
    """(시작, 끝, 문단 xml) 목록. 표 안의 문단도 모두 나온다."""
    return [(m.start(), m.end(), m.group(0)) for m in P_RE.finditer(xml)]


def _matches(p, head_x):
    """문단이 head 로 시작하는가. 앞 공백은 양쪽 다 떼고 견준다.

    매뉴얼 문서는 같은 줄인데도 앞 공백이 다섯 칸이었다 여섯 칸이었다 한다.
    공백까지 맞추라고 하면 작업 목록이 문서의 오타를 따라가야 한다.
    """
    return para_text(p).lstrip().startswith(head_x.lstrip())


def replace_para(xml, head, new, count=None):
    """head 로 시작하는 문단을 new 로 바꾼다. (바꾼 수, 걸린 수, 새 xml)

    count 를 주면 new 는 목록이고 걸린 문단을 차례로 하나씩 맡는다.
    """
    head_x = xml_escape(head)
    news = [xml_escape(v) for v in new] if count else [xml_escape(new)]

    out, last, changed, matched = [], 0, 0, 0
    for s, e, p in para_spans(xml):
        if matched >= len(news) or not _matches(p, head_x):
            continue
        p2, did = rewrite_para(p, news[matched])
        matched += 1
        if did:
            changed += 1
        out.append(xml[last:s])
        out.append(p2)
        last = e
    out.append(xml[last:])
    return changed, matched, "".join(out)


def has_para(xml, text):
    """글자가 똑같은 문단이 이미 있는가."""
    text_x = xml_escape(text)
    return any(para_text(p) == text_x for _, _, p in para_spans(xml))


def text_present(xml, text):
    """그 글자가 문서에 있는가. run 여러 개에 걸쳐 쪼개져 있어도 찾는다.

    '이미 적용되었는가' 를 볼 때 쓴다. 한 <hp:t> 안에서만 찾으면, 넣어 둔 글이
    나중에 여러 run 으로 쪼개진 경우 '못 바꿈' 으로 잘못 보고한다.
    """
    if not text:
        return True
    text_x = xml_escape(text)
    if text_x in xml:
        return True
    return any(text_x in para_text(p) for _, _, p in para_spans(xml))


def clear_para(xml, text):
    """글자가 그것뿐인 문단을 비운다. (비운 수, 새 xml)

    글을 지우고 남은 ' ② ' 같은 번호 조각을 걷어낼 때 쓴다. 문단 전체가 그
    조각일 때만 비우므로, 같은 글자가 다른 줄 한가운데 있어도 다치지 않는다.
    비워 둔 문단은 drop_emptied_paragraphs 가 걷어간다.
    """
    text_x = xml_escape(text).strip()
    out, last, n = [], 0, 0
    for s, e, p in para_spans(xml):
        if para_text(p).strip() != text_x:
            continue
        n += 1
        out.append(xml[last:s])
        out.append(T_RE.sub("<hp:t></hp:t>", p))
        last = e
    out.append(xml[last:])
    return n, "".join(out)


def duplicate_para(xml, anchor, news):
    """anchor 문단을 그대로 복제해 바로 아래에 news 만큼 줄을 만든다.

    표 칸 안의 문단을 복제하므로 같은 칸 안에 줄이 늘어난다. 문단모양·글자모양이
    그대로 따라오므로 새 줄도 위 줄과 똑같이 보인다.
    """
    anchor_x = xml_escape(anchor)
    for s, e, p in para_spans(xml):
        if para_text(p) != anchor_x:
            continue
        added = []
        for v in news:
            p2, _ = rewrite_para(p, xml_escape(v))
            added.append(p2)
        return len(news), xml[:e] + "".join(added) + xml[e:]
    return 0, xml


# ------------------------------------------------------------ 내어쓰기 주기
#
# 항목 문단의 문단모양은 내어쓰기가 0 이다(intent="0", left="0"). 그래서 줄이
# 감기면 다음 줄이 칸 왼쪽 끝으로 붙어, 번호와 글이 뒤엉켜 읽기 어렵다.
#   left = X, intent = -X  → 첫 줄은 0 에서 시작하고 감긴 줄은 X 에서 시작한다.
# X 는 '  ① ' 같은 머리 글자의 너비다. 1pt = 100 HWPUNIT, 전각 한 글자 ≈ 글자 크기.

PARAPR_RE = re.compile(r'<hh:paraPr id="(\d+)"[^>]*>.*?</hh:paraPr>', re.S)
MARGIN_RE = re.compile(r"<hh:margin>.*?</hh:margin>", re.S)


def _strip_id(s):
    return re.sub(r'\sid="\d+"', "", s, count=1)


def clone_parapr(header, base_id, left, first=0):
    """base_id 문단모양을 복제해 들여쓰기·내어쓰기를 준 새 문단모양을 만든다.

    left  ― 감긴 줄이 시작하는 자리
    first ― 첫 줄(항목 번호)이 시작하는 자리. intent = first - left 로 넣는다.
            first=0 이면 예전과 같은 순수 내어쓰기다.

    **<hh:margin> 을 하나만 고치면 안 된다.** 문단모양 안에는 여백이 두 벌 있다.
      <hp:case>    한글 2016 이후가 보는 값
      <hp:default> 그 이름공간을 모르는 판이 보는 값
    예전에는 앞의 것만 고쳐서, default 쪽에 매뉴얼 원본 값(left=0, intent=-5488)이
    그대로 남아 있었다. 그래서 들여쓰기가 글자 앞 공백에만 기대게 되고 판에 따라
    자리가 달라졌다. 두 벌을 똑같이 고친다.

    똑같은 것이 이미 있으면 그것을 쓴다. 다시 돌려도 문단모양이 불어나지 않는다.
    (새 id, 새 header)
    """
    base = None
    ids = []
    for m in PARAPR_RE.finditer(header):
        ids.append(int(m.group(1)))
        if m.group(1) == str(base_id):
            base = m.group(0)
    if base is None:
        return None, header

    def fix(mm):
        s = mm.group(0)
        s = re.sub(r'(<hc:intent value=")-?\d+(")', r"\g<1>%d\g<2>" % (first - left), s)
        s = re.sub(r'(<hc:left value=")-?\d+(")', r"\g<1>%d\g<2>" % left, s)
        return s

    body = MARGIN_RE.sub(fix, base)          # 두 벌 모두
    if body == base:
        return int(base_id), header                  # 이미 그 값이다

    want = _strip_id(body)
    for m in PARAPR_RE.finditer(header):
        if _strip_id(m.group(0)) == want:
            return int(m.group(1)), header           # 같은 것이 이미 있다

    new_id = max(ids) + 1
    body = re.sub(r'id="\d+"', 'id="%d"' % new_id, body, count=1)
    header = header.replace("</hh:paraProperties>", body + "</hh:paraProperties>", 1)
    header = re.sub(r'(<hh:paraProperties itemCnt=")(\d+)(")',
                    lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3),
                    header, count=1)
    return new_id, header


def set_para_indent(xml, header, head, left, first=0):
    """head 로 시작하는 문단들에 들여쓰기·내어쓰기를 준다. (바꾼 수, 새 xml, 새 header)"""
    head_x = xml_escape(head)
    out, last, n = [], 0, 0
    for s, e, p in para_spans(xml):
        if not _matches(p, head_x):
            continue
        m = re.search(r'paraPrIDRef="(\d+)"', p)
        if not m:
            continue
        new_id, header = clone_parapr(header, m.group(1), left, first)
        if new_id is None:
            continue
        n += 1
        out.append(xml[last:s])
        out.append(p[:m.start(1)] + str(new_id) + p[m.end(1):])
        last = e
    out.append(xml[last:])
    return n, "".join(out), header


def strip_leading(xml, head):
    """head 로 시작하는 문단의 앞 공백을 뗀다. (뗀 수, 새 xml)

    들여쓰기를 글자 앞 공백이 아니라 문단모양으로 주려고 할 때 먼저 부른다.
    ("t", …) 로는 안 된다. 그것은 <hp:t> 하나 안에서만 찾는데, 라벨과 값이
    다른 run 으로 쪼개진 줄에서는 '    ∙ {{' 가 두 run 에 걸쳐 있기 때문이다.
    """
    head_x = xml_escape(head)
    out, last, n = [], 0, 0
    for s, e, p in para_spans(xml):
        if not _matches(p, head_x):
            continue
        t = para_text(p)
        # 한 문단 안에 줄바꿈이 들어 있으면 그 뒤도 줄 앞이다. 같이 뗀다.
        # (para_text 는 <hp:lineBreak/> 를 글자 그대로 돌려주므로 그대로 살아 남는다.)
        new = re.sub(r"(<hp:lineBreak/>)\s+", r"\1", t.lstrip())
        if new == t:
            continue
        p2, did = rewrite_para(p, new)
        if not did:
            continue
        out.append(xml[last:s])
        out.append(p2)
        last = e
        n += 1
    out.append(xml[last:])
    return n, "".join(out)


def split_at_linebreak(xml, head):
    """head 로 시작하는 문단을 문단 안 줄바꿈에서 갈라 여러 문단으로 만든다.

    한 문단 안의 줄바꿈은 '이어지는 줄' 이라 내어쓰기(left)만큼 더 들어간다.
    항목이 두 줄로 되어 있는데 둘째 줄만 깊이 들어가 보이면 이것이 원인이다.
    문단을 나누면 두 줄 다 첫 줄이 되어 자리가 같아진다.
    """
    out, last, n = [], 0, 0
    for s, e, p in para_spans(xml):
        if not _matches(p, xml_escape(head)):
            continue
        parts = [t for t in para_text(p).split("<hp:lineBreak/>") if t.strip()]
        if len(parts) < 2:
            continue
        made = []
        for t in parts:
            p2, _ = rewrite_para(p, t.strip())
            made.append(LINESEG_RE.sub("", p2))
        out.append(xml[last:s])
        out.append("".join(made))
        last = e
        n += 1
    out.append(xml[last:])
    return n, "".join(out)


LINESEG_RE = re.compile(r"<hp:linesegarray>.*?</hp:linesegarray>", re.S)


def move_paras(xml, heads, anchor):
    """heads 로 시작하는 문단들을 잘라 anchor 문단 바로 뒤로 옮긴다.

    표 칸을 건너뛰어 옮길 수 있다(문단 XML 은 그 자체로 완결되어 있고, 문단모양·
    글자모양은 header 를 가리키므로 다른 칸에 붙여도 그대로 보인다).

    옮긴 문단의 <hp:linesegarray> 는 지운다. 이것은 '이 문단이 몇 줄이었는지' 를
    적어 둔 캐시라, 폭이 다른 칸으로 옮기고도 남겨 두면 옛 줄 수로 그려져 글자가
    겹친다.

    **이미 anchor 와 같은 칸에서 바로 뒤에 붙어 있으면 아무것도 하지 않는다.**
    두 덩어리를 맞바꾸는 데 쓰므로, 이 검사가 없으면 다시 돌릴 때마다 자리가
    엎치락뒤치락한다. 문서 차례만 보면 안 된다 ― 왼쪽 칸의 마지막 문단과 오른쪽
    칸의 첫 문단은 차례로는 이웃이라, 칸을 안 따지면 '이미 제자리' 로 잘못 본다.
    그러다 두 덩어리가 한 칸에 몰리고 다른 칸이 텅 비면, **문단이 하나도 없는 칸을
    한글이 열지 못해 죽는다**(엑셀에는 462 런타임 오류로 올라온다).
    """
    cells = sorted(all_spans(xml, "hp:tc"))

    def cell_of(pos):
        """그 자리를 품은 가장 안쪽 칸. 표 밖이면 None."""
        best = None
        for s, e in cells:
            if s <= pos < e and (best is None or s > best[0]):
                best = (s, e)
        return best

    spans = list(para_spans(xml))
    texts = [para_text(p).strip() for _, _, p in spans]

    a, idx = None, {}
    for i, t in enumerate(texts):
        for h in heads:
            hx = xml_escape(h).strip()
            if h not in idx and t.startswith(hx):
                idx[h] = i
    anchor_x = xml_escape(anchor).strip()
    for i, t in enumerate(texts):
        if i not in idx.values() and t.startswith(anchor_x):
            a = i
            break
    if a is None or len(idx) != len(heads):
        return 0, xml

    order = [idx[h] for h in heads]
    home = cell_of(spans[a][0])
    if (order == list(range(a + 1, a + 1 + len(heads))) and
            all(cell_of(spans[i][0]) == home for i in order)):
        return 0, xml                                  # 이미 제자리

    # 옮기고 나면 원래 칸이 텅 비는가. 비면 한글이 문서를 열지 못한다.
    for i in order:
        src = cell_of(spans[i][0])
        if src is None or src == home:
            continue
        left = [j for j, (s, _, _) in enumerate(spans)
                if cell_of(s) == src and j not in order]
        if not left:
            return 0, xml                              # 칸을 비울 수 없다

    MARK = "\x00MOVE\x00"
    marks = sorted([(spans[i][0], spans[i][1], i) for i in order] +
                   [(spans[a][0], spans[a][1], -1)])
    out, last, block = [], 0, {}
    for s, e, i in marks:
        out.append(xml[last:s])
        if i < 0:
            out.append(xml[s:e] + MARK)                # anchor 는 두고 자리만 표시
        else:
            block[i] = LINESEG_RE.sub("", xml[s:e])
        last = e
    out.append(xml[last:])
    return len(order), "".join(out).replace(MARK, "".join(block[i] for i in order), 1)


# --------------------------------------------------------- {{태그}} 파란색으로
#
# 파란 글씨는 '여기는 그대로 두지 말고 확인해서 고치라' 는 뜻으로 쓰인다.
# 태그가 든 run 만 골라 글자색을 바꾸고, 앞의 라벨은 검은색으로 남긴다.

CHARPR_RE = re.compile(r'<hh:charPr id="(\d+)"[^>]*>.*?</hh:charPr>', re.S)
TAG_RE = re.compile(r"\{\{[^{}]*\}\}")
RUN_OPEN_RE = re.compile(r'<hp:run charPrIDRef="(\d+)"[^>]*>')


def clone_charpr(header, base_id, color):
    """base_id 글자모양을 복제해 색만 바꾼 새 글자모양을 만든다. (새 id, 새 header)"""
    base, ids = None, []
    for m in CHARPR_RE.finditer(header):
        ids.append(int(m.group(1)))
        if m.group(1) == str(base_id):
            base = m.group(0)
    if base is None:
        return None, header

    body = re.sub(r'(textColor=")[^"]*(")', r"\g<1>%s\g<2>" % color, base, count=1)
    if body == base:
        return int(base_id), header                    # 이미 그 색이다

    want = _strip_id(body)
    for m in CHARPR_RE.finditer(header):
        if _strip_id(m.group(0)) == want:
            return int(m.group(1)), header

    new_id = max(ids) + 1
    body = re.sub(r'id="\d+"', 'id="%d"' % new_id, body, count=1)
    header = header.replace("</hh:charProperties>", body + "</hh:charProperties>", 1)
    header = re.sub(r'(<hh:charProperties itemCnt=")(\d+)(")',
                    lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3),
                    header, count=1)
    return new_id, header


def color_tag_runs(xml, header, color):
    """{{태그}} 가 든 run 을 색깔 run 으로 가른다. (바꾼 수, 새 xml, 새 header)

    run 하나에 <hp:t> 가 하나뿐일 때만 가른다. 줄바꿈이나 개체가 섞인 run 은
    가르다가 깨질 수 있으므로 통째로 색만 바꾼다.
    """
    out, last, n = [], 0, 0
    for m in RUN_OPEN_RE.finditer(xml):
        if m.start() < last:
            continue                                   # 이미 지나간 자리
        close = xml.find("</hp:run>", m.end())
        if close < 0:
            continue
        inner = xml[m.end():close]
        if "<hp:run" in inner:
            continue                                   # 표를 품은 run 은 건너뛴다
        ts = list(T_RE.finditer(inner))
        if len(ts) != 1:
            continue
        body = ts[0].group(1)
        tags = list(TAG_RE.finditer(body))
        if not tags and "{{" not in body and "}}" not in body:
            continue
        new_id, header = clone_charpr(header, m.group(1), color)
        if new_id is None or str(new_id) == m.group(1):
            continue                                   # 이미 그 색이다

        pre, post = inner[:ts[0].start()], inner[ts[0].end():]
        if not tags:
            # 태그가 두 run 으로 쪼개져 '{{학교' 나 '명}}' 만 든 run 이다.
            # 가를 자리가 없으니 run 째로 파랗게 한다.
            pieces = [(str(new_id), body)]
        else:
            pieces, pos = [], 0
            for t in tags:
                if t.start() > pos:
                    pieces.append((m.group(1), body[pos:t.start()]))
                pieces.append((str(new_id), t.group(0)))
                pos = t.end()
            if pos < len(body):
                pieces.append((m.group(1), body[pos:]))

        runs = []
        for i, (cid, seg) in enumerate(pieces):
            open_tag = re.sub(r'charPrIDRef="\d+"', 'charPrIDRef="%s"' % cid,
                              m.group(0), count=1)
            runs.append(open_tag +
                        (pre if i == 0 else "") +
                        "<hp:t>" + seg + "</hp:t>" +
                        (post if i == len(pieces) - 1 else "") +
                        "</hp:run>")
        out.append(xml[last:m.start()])
        out.append("".join(runs))
        last = close + len("</hp:run>")
        n += 1
    out.append(xml[last:])
    return n, "".join(out), header


def all_spans(xml, name):
    """<name …> … </name> 의 (시작, 끝) 을 **겹친 것까지 모두** 돌려준다.

    표 안에 표가 들어 있으므로 바깥쪽만 세면 안쪽 칸을 놓친다.
    """
    open_re = re.compile(r"<%s\b[^>]*>" % name)
    events = []
    for m in open_re.finditer(xml):
        events.append((m.start(), 0 if m.group(0).endswith("/>") else 1, m.end()))
    for m in re.finditer(r"</%s>" % name, xml):
        events.append((m.start(), -1, m.end()))
    events.sort()

    stack, out = [], []
    for pos, kind, end in events:
        if kind == 0:
            out.append((pos, end))
        elif kind == 1:
            stack.append(pos)
        elif stack:
            out.append((stack.pop(), end))
    return sorted(out)


def innermost(spans, pos):
    """pos 를 품은 것들 중 가장 안쪽(시작이 가장 늦은) 것."""
    best = None
    for s, e in spans:
        if s < pos < e and (best is None or s > best[0]):
            best = (s, e)
    return best


CELLADDR_RE = re.compile(r'<hp:cellAddr colAddr="(\d+)" rowAddr="(\d+)"/>')
CELLSPAN_RE = re.compile(r'<hp:cellSpan colSpan="(\d+)" rowSpan="(\d+)"/>')
# 빈 칸의 문단은 글자 없는 run 하나로 되어 있다. 두 가지 꼴이 다 나온다.
EMPTY_RUN_RE = re.compile(
    r'<hp:run charPrIDRef="(\d+)"\s*/>|<hp:run charPrIDRef="(\d+)"><hp:t></hp:t></hp:run>')


def find_label_cell(xml, label, cells):
    """칸 글자가 label 과 똑같은 첫 <hp:tc> 의 (시작, 끝). 없으면 None."""
    label_x = xml_escape(label)
    for m in re.finditer(r"<hp:t>%s</hp:t>" % re.escape(label_x), xml):
        cell = innermost(cells, m.start())
        if not cell:
            continue
        # 칸 안의 글자가 라벨 하나뿐일 때만 라벨 칸으로 본다.
        # (본문에 같은 낱말이 섞여 있어도 걸리지 않는다)
        body = "".join(T_RE.findall(xml[cell[0]:cell[1]]))
        if body.strip() == label_x.strip():
            return cell
    return None


def fill_neighbour_cell(xml, label, tag, direction):
    """라벨 칸의 오른쪽(right)/아래(below) 칸에 태그를 써 넣는다. (성공?, xml, 메모)"""
    cells = all_spans(xml, "hp:tc")
    tbls = all_spans(xml, "hp:tbl")

    found = find_label_cell(xml, label, cells)
    if not found:
        return False, xml, "라벨 칸 '%s' 없음" % label
    s, e = found

    addr = CELLADDR_RE.search(xml[s:e])
    span = CELLSPAN_RE.search(xml[s:e])
    if not addr:
        return False, xml, "'%s' 칸 주소 없음" % label
    col, row = int(addr.group(1)), int(addr.group(2))
    cspan = int(span.group(1)) if span else 1
    rspan = int(span.group(2)) if span else 1
    want = (col + cspan, row) if direction == "right" else (col, row + rspan)

    # 라벨 칸이 든 표(가장 안쪽 표) 안에서만 찾는다.
    tbl = innermost(tbls, s)
    if tbl is None:
        return False, xml, "'%s' 이 든 표를 못 찾음" % label

    target = None
    for cs, ce in cells:
        if not (tbl[0] < cs and ce < tbl[1]):
            continue
        if innermost(tbls, cs) != tbl:      # 안쪽 표의 칸은 건너뛴다
            continue
        a = CELLADDR_RE.search(xml[cs:ce])
        if a and (int(a.group(1)), int(a.group(2))) == want:
            target = (cs, ce)
            break
    if target is None:
        return False, xml, "'%s' %s 칸(%d,%d) 없음" % (label, direction, want[0], want[1])

    cs, ce = target
    cell = xml[cs:ce]
    m = EMPTY_RUN_RE.search(cell)
    if not m:
        body = "".join(T_RE.findall(cell)).strip()
        return False, xml, "'%s' %s 칸이 비어 있지 않음(%r)" % (label, direction, body[:20])

    char_id = m.group(1) or m.group(2)
    filled = ('<hp:run charPrIDRef="%s"><hp:t>%s</hp:t></hp:run>'
              % (char_id, xml_escape(tag)))
    cell = cell[:m.start()] + filled + cell[m.end():]
    return True, xml[:cs] + cell + xml[ce:], ""


# ------------------------------------------------------------------ 본체

KEEP_TAG_RE = re.compile(
    r"<hp:(tbl|pic|secPr|colPr|ctrl|container|rect|line|ellipse|polygon|arc|ole|"
    r"chart|equation|textart|connectLine|curve|video|formCtrl)\b")


def drop_emptied_paragraphs(before, after):
    """이번에 글을 지워 빈 껍데기가 된 문단을 걷어낸다.

    글만 지우고 문단을 남기면, 번호가 매겨진 문단에서는 '2.' 같은 번호만 덩그러니
    남는다. 이번에 비운 문단만 골라 지우므로, 원래부터 줄 간격용으로 비어 있던
    문단은 건드리지 않는다.
    """
    b_spans = all_spans(before, "hp:p")
    a_spans = all_spans(after, "hp:p")
    if len(b_spans) != len(a_spans):
        return after            # 문단 수가 달라졌으면 손대지 않는다

    cells = all_spans(after, "hp:tc")
    drop = []
    for (bs, be), (as_, ae) in zip(b_spans, a_spans):
        had = "".join(T_RE.findall(before[bs:be])).strip()
        now = "".join(T_RE.findall(after[as_:ae])).strip()
        if not had or now:
            continue
        if KEEP_TAG_RE.search(after[as_:ae]):
            continue
        cell = innermost(cells, as_)
        if cell is None:
            continue            # 표 밖의 문단은 쪽 배치에 영향이 커서 두고 본다
        # 칸에 문단이 하나뿐이면 지우지 않는다(빈 칸이 되어 표가 무너진다)
        siblings = [1 for s, e in a_spans if cell[0] < s and e < cell[1]]
        if len(siblings) <= 1:
            continue
        drop.append((as_, ae))

    if not drop:
        return after
    out, cursor = [], 0
    for s, e in sorted(drop):
        out.append(after[cursor:s])
        cursor = e
    out.append(after[cursor:])
    return "".join(out)


# 문단을 늘리거나 문단모양을 바꾸는 작업. 글자를 고치는 작업이 다 끝난 뒤에
# 따로 돌린다. drop_emptied_paragraphs 가 '문단 수가 그대로일 것' 을 전제로
# 하기 때문이다(수가 달라지면 손을 놓아 버려서, 비운 문단이 그대로 남는다).
STRUCTURAL = ("dup", "indent")


def apply_ops(xml, ops, header=None):
    """작업을 차례로 적용한다. (새 xml, 새 header, 끝난 수, 이미 된 수, 못한 것)

    이미 심어 둔 파일에 다시 돌려도 안전하도록, 옛 글자가 없고 새 글자가 이미
    있으면 '이미 적용됨' 으로 본다.
    """
    notes, done, already = [], 0, 0

    def has(text):
        return text_present(xml, text)

    for op in ops:
        kind = op[0]
        if kind == "clearpara":
            n, xml = clear_para(xml, op[1])
            if n:
                done += 1
            else:
                already += 1        # 이미 지워졌거나 원래 없던 조각

        elif kind == "para":
            n, hit, xml = replace_para(xml, op[1], op[2])
            if hit:
                done += 1
            elif has_para(xml, op[2]):
                already += 1
            else:
                notes.append("문단 못 찾음: %r" % op[1][:28])

        elif kind == "paraseq":
            n, hit, xml = replace_para(xml, op[1], op[2], count=len(op[2]))
            if hit == len(op[2]):
                done += 1
            elif all(has_para(xml, v) for v in op[2]):
                already += 1
            else:
                notes.append("%r 문단을 %d개 중 %d개만 찾음"
                             % (op[1], len(op[2]), hit))

        elif kind == "dup":
            if all(has_para(xml, v) for v in op[2]):
                already += 1
            else:
                n, xml = duplicate_para(xml, op[1], op[2])
                if n:
                    done += 1
                else:
                    notes.append("복제할 문단 없음: %r" % op[1][:28])

        elif kind == "indent":
            first = op[3] if len(op) > 3 else 0
            n, xml, header = set_para_indent(xml, header, op[1], op[2], first)
            if n:
                done += 1
            else:
                notes.append("내어쓰기 줄 없음: %r" % op[1][:28])

        elif kind == "splitbr":
            n, xml = split_at_linebreak(xml, op[1])
            if n:
                done += 1
            else:
                already += 1        # 이미 갈라져 있다

        elif kind == "striplead":
            n, xml = strip_leading(xml, op[1])
            if n:
                done += 1
            else:
                already += 1        # 이미 공백이 없다

        elif kind == "movepara":
            n, xml = move_paras(xml, op[1], op[2])
            if n:
                done += 1
            else:
                already += 1        # 이미 제자리이거나 옮길 문단이 없다

        elif kind == "tagcolor":
            n, xml, header = color_tag_runs(xml, header, op[1])
            if n:
                done += 1
            else:
                already += 1        # 이미 다 그 색이다

        elif kind == "t":
            n, xml = replace_text(xml, op[1], op[2])
            if n:
                done += 1
            elif has(op[2]):
                already += 1
            else:
                notes.append("못 바꿈: %r" % op[1][:28])
        elif kind == "run":
            n, xml = replace_whole_run(xml, op[1], op[2])
            if n:
                done += 1
            else:
                already += 1        # 이미 지워졌거나 원래 없던 조각

        elif kind == "seq":
            n, xml = replace_sequence(xml, op[1], op[2])
            if n == len(op[2]):
                done += 1
            elif n == 0 and all(has(v) for v in op[2]):
                already += 1
            else:
                notes.append("%r 를 %d개 중 %d개만 바꿈" % (op[1], len(op[2]), n))
        else:
            ok, xml, why = fill_neighbour_cell(xml, op[1], op[2], kind)
            if ok:
                done += 1
            elif has(op[2]):
                already += 1
            else:
                notes.append(why)
    return xml, header, done, already, notes


def main():
    dry = "--dry-run" in sys.argv
    ok_all = True

    for name in sorted(OPS):
        path = os.path.join(TDIR, name)
        if not os.path.exists(path):
            print(f"  {name:<28} ★ 파일 없음")
            ok_all = False
            continue

        entries, info = read_section(path)
        sec, hdr = "Contents/section0.xml", "Contents/header.xml"
        xml = entries[sec].decode("utf-8")
        header = entries[hdr].decode("utf-8")

        before, header_before = xml, header
        ops = OPS[name]

        # ① 글자 고치기 → ② 비워진 문단 걷어내기 → ③ 문단 늘리기·문단모양 바꾸기
        #    ②는 문단 수가 그대로여야 도는 검사라 ③보다 먼저 해야 한다.
        text_ops = [o for o in ops if o[0] not in STRUCTURAL]
        struct_ops = [o for o in ops if o[0] in STRUCTURAL]

        xml, header, done, already, notes = apply_ops(xml, text_ops, header)
        xml = drop_emptied_paragraphs(before, xml)
        xml, header, d2, a2, n2 = apply_ops(xml, struct_ops, header)
        done, already, notes = done + d2, already + a2, notes + n2

        # 문단마다 '이 글이 몇 줄이었는지' 가 캐시되어 있다(linesegarray).
        # 태그를 넣어 글이 길어지면 한글이 옛 줄 수를 그대로 써서 글자가 겹쳐 찍힌다.
        # 지워 두면 열 때 다시 계산한다.
        xml = re.sub(r"<hp:linesegarray>.*?</hp:linesegarray>", "", xml, flags=re.S)

        if not dry and (xml != before or header != header_before):
            entries[sec] = xml.encode("utf-8")
            entries[hdr] = header.encode("utf-8")
            write_section(path, entries, info)

        tags = sorted(set(re.findall(r"\{\{[^}]*\}\}", "".join(T_RE.findall(xml)))))
        good = done + already
        mark = ("OK" + (f" (이미 {already})" if already else "")
                if good == len(ops) else "★ " + " / ".join(notes))
        print(f"  {name:<28} {good}/{len(ops)}  {mark}")
        print(f"      태그 {len(tags)}: {' '.join(tags)}")
        if good != len(ops):
            ok_all = False

    if dry:
        print("\n(--dry-run: 파일은 고치지 않았습니다)")
    return 0 if ok_all else 1


if __name__ == "__main__":
    print("치환 태그 심기 (신 서식):")
    sys.exit(main())
