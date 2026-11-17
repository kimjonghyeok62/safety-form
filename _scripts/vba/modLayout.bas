Attribute VB_Name = "modLayout"
Option Explicit

'==============================================================================
' 시트 레이아웃 상수 ― 행/열 번호는 반드시 이 모듈을 통해서만 참조한다.
' (과거 출력관리와 입력정보의 행 어긋남으로 엉뚱한 계약이 출력되는 버그가 있었다)
'
' build_workbook.py 의 같은 이름 상수와 반드시 짝이 맞아야 한다.
'==============================================================================

' --- 시트 이름 -----------------------------------------------------------
Public Const SH_INPUT  As String = "입력정보"
Public Const SH_OUTPUT As String = "출력관리"
Public Const SH_CONFIG As String = "설정"
Public Const SH_DICT   As String = "계약유형사전"
Public Const SH_LIB    As String = "ContentLibrary"

' --- 입력정보: ① 기관 기본정보 (3행 머리글 / 4행 값) ---------------------
'
' 위쪽 세 블록은 아래 계약정보와 열 격자를 함께 쓴다. 계약정보에 맞춘 열 너비로는
' 주소·전화번호가 잘리므로 칸을 여러 열에 걸쳐 병합했다. 값은 늘 병합된 칸의
' **첫 열**에 들어간다. 아래 열 번호는 build_workbook.py 의 *_FIELDS 와 짝이다.
Public Const IN_SCHOOL_ROW As Long = 4
Public Const IC_기관구분   As Long = 1          ' A:B  학교 / 교육지원청
Public Const IC_기관명     As Long = 3          ' C:D
Public Const IC_고유번호   As Long = 5          ' E:F
Public Const IC_주소       As Long = 7          ' G:I
Public Const IC_전화번호   As Long = 10         ' J:K
Public Const IC_팩스번호   As Long = 12         ' L:M

' --- 입력정보: ② 결재선·기관대표자 (5행 머리글 / 6행 값) -----------------
' 결재선의 부서장과 기관대표자는 다르다.
'   학교      부서장 = 교장,  기관대표자 = 교장   (같다)
'   교육지원청 부서장 = 과장,  기관대표자 = 교육장 (다르다)
Public Const IN_SIGN_ROW     As Long = 6
Public Const SC_담당자직위   As Long = 1        ' A:B
Public Const SC_담당자성명   As Long = 3        ' C
Public Const SC_담당자전화   As Long = 4        ' D:E
Public Const SC_관리자직위   As Long = 6        ' F:G
Public Const SC_관리자성명   As Long = 8        ' H
Public Const SC_부서장직위   As Long = 9        ' I:J
Public Const SC_부서장성명   As Long = 11       ' K:L
Public Const SC_대표자직위   As Long = 13       ' M:N  기관대표자
Public Const SC_대표자성명   As Long = 15       ' O:Q

' --- 입력정보: ③ 비상연락망·관할기관 (8행 머리글 / 9행 값) ---------------
Public Const IN_EMERG_ROW  As Long = 9
Public Const EC_소방서     As Long = 1          ' A:B
Public Const EC_소방서전화 As Long = 3          ' C:D
Public Const EC_병원       As Long = 5          ' E:F
Public Const EC_병원전화   As Long = 7          ' G:H
Public Const EC_노동지청   As Long = 9          ' I:K
Public Const EC_노동지청전화 As Long = 12       ' L:M

' --- 입력정보: ④ 계약정보 (11행 머리글 / 12~111행 값) --------------------
' 구역 사이에 빈 줄은 두지 않는다. 계약정보가 화면에 한 줄이라도 더 들어오도록
' 없앴고, 구역은 제목 줄 윗변의 가로줄로 가른다.
Public Const IN_HEAD_ROW  As Long = 11
Public Const IN_FIRST_ROW As Long = 12
Public Const IN_LAST_ROW  As Long = 111         ' 계약 100건

Public Const CC_연번         As Long = 1        ' A
Public Const CC_계약명       As Long = 2        ' B
Public Const CC_시작일       As Long = 3        ' C
Public Const CC_종료일       As Long = 4        ' D
Public Const CC_점검요일     As Long = 5        ' E  계약기간 바로 옆
Public Const CC_계약금액     As Long = 6        ' F
Public Const CC_업체명       As Long = 7        ' G
Public Const CC_사업자번호   As Long = 8        ' H
Public Const CC_업체주소     As Long = 9        ' I
Public Const CC_업체대표     As Long = 10       ' J
Public Const CC_대표자전화   As Long = 11       ' K
Public Const CC_안전담당자   As Long = 12       ' L  업체 담당자 (직위 성명)
Public Const CC_담당자전화   As Long = 13       ' M
Public Const CC_고소         As Long = 14       ' N
Public Const CC_화기         As Long = 15       ' O
Public Const CC_밀폐         As Long = 16       ' P
Public Const CC_전기         As Long = 17       ' Q
Public Const CC_계획서유형   As Long = 18       ' R
Public Const CC_계약유형     As Long = 19       ' S  계약유형사전 자동 매칭 결과
Public Const CC_비고         As Long = 20       ' T  진행 / 완료
Public Const CC_출력버튼     As Long = 21       ' U  ▶ 출력관리 버튼 자리
' V 는 숨김이다. 비고를 사용자가 직접 골랐다는 표시("수동")만 담는다.
' 줄 전체가 진행/완료 색으로 덮이므로 색으로는 표시할 수가 없어 칸을 따로 두었다.
Public Const CC_비고수동     As Long = 22       ' V (숨김)

Public Const WORK_FIRST_COL As Long = CC_고소   ' 작업유형 4개 열의 시작
Public Const WORK_LAST_COL  As Long = CC_전기

' --- 출력관리 ------------------------------------------------------------
Public Const OUT_TITLE_ROW  As Long = 1         ' 2행은 제목 아래 빈 줄
Public Const OUT_BTN_ROW    As Long = 3         ' 버튼줄 (한 줄에 8개)
Public Const OUT_HEAD_ROW   As Long = 4
Public Const OUT_RULE_ROW   As Long = 5         ' 판정기준 안내 행
Public Const OUT_FIRST_ROW  As Long = 6
Public Const OUT_LAST_ROW   As Long = 105       ' 6~105 = 100건

' 출력관리 행 = 입력정보 행 + OUT_ROW_OFFSET
Public Const OUT_ROW_OFFSET As Long = -6

Public Const OC_선택     As Long = 1            ' A  체크박스
Public Const OC_연번     As Long = 2            ' B
Public Const OC_계약명   As Long = 3            ' C
Public Const OC_계약기간 As Long = 4            ' D  '365일' ― 입력정보 날짜에서 계산
Public Const OC_점검요일 As Long = 5            ' E  여기서 바꾸면 입력정보로 되돌아간다
Public Const OC_서식1    As Long = 6            ' F ~ T = 서식 1~15
Public Const OC_서식15   As Long = 20
Public Const OC_요약     As Long = 21           ' U

Public Const FORM_COUNT       As Long = 15
' [1~7번 출력] 이 묶는 범위. 8번(안전보건협의체 회의록)은 월 1회 반복이라
' 순회·합동처럼 자기 버튼(Btn_ExportCouncil)으로 따로 나간다.
Public Const FORM_MERGE_LAST  As Long = 7

' --- 결재선 보관함 (ContentLibrary 시트 AA열) ----------------------------
' 기관구분을 잘못 골랐다가 되돌렸을 때 적어 둔 값을 잃지 않게 숨겨 둔다.
'   AA1        = 지금 적용 중인 기관구분 (바뀌기 '전' 값을 알아내는 유일한 방법)
'   AA2 이하   = AA 기관구분 이름 / AB~AF 기관정보 5칸 / AG~AK 결재선 성명·전화 5칸
Public Const STASH_COL       As Long = 27       ' AA
Public Const STASH_STATE_ROW As Long = 1
Public Const STASH_FIRST_ROW As Long = 2
Public Const STASH_LAST_ROW  As Long = 20       ' 기관구분은 두 가지뿐이라 넉넉하다

' 서식 번호(1~15) → 출력관리 열 번호
Public Function FormCol(ByVal formNo As Long) As Long
    FormCol = OC_서식1 + formNo - 1
End Function

' 입력정보 행 → 출력관리 행
Public Function OutRow(ByVal inputRow As Long) As Long
    OutRow = inputRow + OUT_ROW_OFFSET
End Function

' 출력관리 행 → 입력정보 행
Public Function InputRow(ByVal outputRow As Long) As Long
    InputRow = outputRow - OUT_ROW_OFFSET
End Function

'==============================================================================
' 공용 헬퍼
'==============================================================================

Public Function WsInput() As Worksheet
    Set WsInput = ThisWorkbook.Worksheets(SH_INPUT)
End Function

Public Function WsOutput() As Worksheet
    Set WsOutput = ThisWorkbook.Worksheets(SH_OUTPUT)
End Function

Public Function WsConfig() As Worksheet
    Set WsConfig = ThisWorkbook.Worksheets(SH_CONFIG)
End Function

Public Function WsDict() As Worksheet
    On Error Resume Next
    Set WsDict = ThisWorkbook.Worksheets(SH_DICT)
    On Error GoTo 0
End Function

' 계약이 실제로 입력된 행인지 (계약명이 있어야 유효)
Public Function IsContractRow(ByVal inputRow As Long) As Boolean
    If inputRow < IN_FIRST_ROW Or inputRow > IN_LAST_ROW Then Exit Function
    IsContractRow = (Trim$(CStr(WsInput().Cells(inputRow, CC_계약명).Value)) <> "")
End Function

' 출력관리에서 체크된 계약의 입력정보 행 번호 목록.
'
' 하나도 체크되어 있지 않으면 **커서가 놓인 줄**을 그 계약으로 본다.
' 체크를 깜빡했다고 경고창을 띄우면 [확인] → [체크] → [버튼] 으로 세 번을 더 눌러야 한다.
Public Function CheckedInputRows() As Variant
    Dim ws As Worksheet: Set ws = WsOutput()
    Dim acc() As Long, n As Long
    ReDim acc(0 To OUT_LAST_ROW - OUT_FIRST_ROW)
    n = 0

    Dim r As Long
    For r = OUT_FIRST_ROW To OUT_LAST_ROW
        If ws.Cells(r, OC_선택).Value = True Then
            If IsContractRow(InputRow(r)) Then
                acc(n) = InputRow(r)
                n = n + 1
            End If
        End If
    Next r

    If n = 0 Then
        Dim cur As Long
        cur = CursorContractRow()
        If cur > 0 Then
            acc(0) = cur
            n = 1
        End If
    End If

    If n = 0 Then
        CheckedInputRows = Array()
    Else
        ReDim Preserve acc(0 To n - 1)
        CheckedInputRows = acc
    End If
End Function

' 출력관리에서 커서가 놓인 줄의 입력정보 행. 계약이 없는 줄이면 0.
Public Function CursorContractRow() As Long
    On Error Resume Next
    Dim ws As Worksheet: Set ws = WsOutput()
    If ActiveSheet Is Nothing Then Exit Function
    If ActiveSheet.Name <> SH_OUTPUT Then Exit Function

    Dim r As Long
    r = ActiveCell.Row
    If r < OUT_FIRST_ROW Or r > OUT_LAST_ROW Then Exit Function
    If IsContractRow(InputRow(r)) Then CursorContractRow = InputRow(r)
    On Error GoTo 0
End Function

' 계약기간(일수). 시작/종료가 비었거나 뒤집혀 있으면 0.
Public Function ContractDays(ByVal inputRow As Long) As Long
    Dim ws As Worksheet: Set ws = WsInput()
    Dim s As Variant, e As Variant
    s = ws.Cells(inputRow, CC_시작일).Value
    e = ws.Cells(inputRow, CC_종료일).Value
    If Not IsDate(s) Or Not IsDate(e) Then Exit Function
    If CDate(e) < CDate(s) Then Exit Function
    ContractDays = DateDiff("d", CDate(s), CDate(e)) + 1
End Function

' 해당 계약에 ○ 로 표시된 작업유형 이름들 ("고소작업|화기작업" 꼴)
Public Function ActiveWorkTypes(ByVal inputRow As Long) As String
    Dim ws As Worksheet: Set ws = WsInput()
    Dim names As Variant
    names = Array("고소작업", "화기작업", "밀폐공간작업", "전기작업")

    Dim acc As String, c As Long
    For c = WORK_FIRST_COL To WORK_LAST_COL
        If IsMarked(ws.Cells(inputRow, c).Value) Then
            If acc <> "" Then acc = acc & "|"
            acc = acc & names(c - WORK_FIRST_COL)
        End If
    Next c
    ActiveWorkTypes = acc
End Function

' ○ 표시 여부 판정. × 와 빈칸은 '아님'. 과거 TRUE/FALSE 값도 함께 받아준다.
Public Function IsMarked(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then
        IsMarked = CBool(v)
    Else
        Select Case Trim$(CStr(v))
            Case "○", "O", "o", "V", "v", "√", "Y", "y": IsMarked = True
            Case Else: IsMarked = False
        End Select
    End If
End Function

' "직위 성명" 을 만든다. 한쪽이 비면 있는 쪽만 돌려준다.
Public Function JoinTitleName(ByVal title As String, ByVal nm As String) As String
    title = Trim$(title)
    nm = Trim$(nm)
    If title <> "" And nm <> "" Then
        JoinTitleName = title & " " & nm
    Else
        JoinTitleName = title & nm
    End If
End Function
