Attribute VB_Name = "modConfig"
Option Explicit

'==============================================================================
' '설정' 시트 읽기. 모든 접근은 이름 정의(Named Range)를 통해서 하므로
' 사용자가 설정 시트에서 행을 밀어 넣어도 코드가 깨지지 않는다.
' 이름이 없거나 값이 비면 코드 안의 기본값으로 되돌아간다.
'==============================================================================

' 색이 곧 상태다. 눈으로 바로 보이므로 사용자가 따로 기억할 필요가 없다.
' (VBA 의 Interior.Color 는 BGR 정수다 ― R + G*256 + B*65536)
Public Const AUTO_FILL_COLOR   As Long = 14744063   ' RGB(255,249,224) 연노랑 ― 자동으로 채운 칸
Public Const MANUAL_FILL_COLOR As Long = 15657703   ' RGB(231,234,238) 연회색 ― 직접 정한 칸
Public Const STRIPE_COLOR      As Long = 16579578   ' RGB(250,251,252) 줄무늬
Public Const INPUT_BG_COLOR    As Long = 16121343   ' RGB(255,253,245) 사람이 채우는 칸
                                                    ' (build_workbook.py 의 INPUT_BG 와 같아야 한다)

' 방금 옮겨 온 줄을 5초간 알리는 색. **진행 중 계약의 연한 초록(#E2F3DB)보다
' 확실히 진해야 한다.** 예전 값 RGB(227,242,217) 은 그 색과 거의 같아서,
' 진행 중인 계약으로 옮겨 가면 강조가 아예 보이지 않았다.
Public Const HIGHLIGHT_COLOR   As Long = 11593670   ' RGB(198,231,176) 연두

' 계약유형 사전 시트의 열 번호.
' 모듈 수준 선언은 반드시 첫 프로시저보다 위에 있어야 한다. 중간에 두면 그 지점부터
' 컴파일이 깨지는데, VBA 는 그것을 실행할 때가 되어서야 알리므로 원인을 찾기 어렵다.
Public Const DICT_유형명   As Long = 1
Public Const DICT_키워드   As Long = 2
Public Const DICT_작업내용 As Long = 3
Public Const DICT_위험요인 As Long = 4
Public Const DICT_안전대책 As Long = 5
Public Const DICT_협의안건 As Long = 6
Public Const DICT_점검착안 As Long = 7
Public Const DICT_작업유형 As Long = 8
' I~O ― 계획서·회의록의 파란 예시 글을 대신할 문구
Public Const DICT_작업절차 As Long = 9
Public Const DICT_점검계획 As Long = 10
Public Const DICT_보호구   As Long = 11
Public Const DICT_교육계획 As Long = 12
Public Const DICT_기계기구 As Long = 13
Public Const DICT_위험물질 As Long = 14
Public Const DICT_대피방법 As Long = 15

' 학교 현장 항목 ― 계획서의 관리 책임란과 회의록 의결란에 들어간다
Public Const DICT_권장시기 As Long = 16
Public Const DICT_학교확인 As Long = 17
Public Const DICT_업체조치 As Long = 18
Public Const DICT_위험등급 As Long = 19
Public Const DICT_합의문   As Long = 20

' 시트에서 떠 온 것을 잠깐 들고 있는 자리. 쓰는 곳은 아래 '계약유형 사전' 절과
' TemplateFolder / TempFolder 이고, ResetLookupCache 가 비운다.
' (선언은 첫 프로시저보다 위여야 하므로 쓰는 곳에서 멀지만 여기 둔다 ― 위 주석 참고)
Private DictCache As Variant
Private DictIndex As Object              ' 유형명 → 사전 안에서의 줄 번호
Private CachedTemplateFolder As String
Private CachedTempFolder As String
Private HolidayCache As Object           ' "yyyymmdd" → True

'--- 단일 값 ------------------------------------------------------------------

Private Function NamedValue(ByVal nm As String, ByVal fallback As Variant) As Variant
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then
        NamedValue = fallback
    ElseIf Trim$(CStr(rng.Cells(1, 1).Value)) = "" Then
        NamedValue = fallback
    Else
        NamedValue = rng.Cells(1, 1).Value
    End If
End Function

Private Function NamedRange(ByVal nm As String) As Range
    On Error Resume Next
    Set NamedRange = ThisWorkbook.Names(nm).RefersToRange
    On Error GoTo 0
End Function

'--- 서식 필요 판정 기준(일수) ------------------------------------------------

Public Function CriteriaPlan() As Long
    CriteriaPlan = CLng(NamedValue("기준_계획서", 2))
End Function

Public Function CriteriaLevel() As Long
    CriteriaLevel = CLng(NamedValue("기준_수준", 2))
End Function

Public Function CriteriaCouncil() As Long
    CriteriaCouncil = CLng(NamedValue("기준_협의체", 30))
End Function

Public Function CriteriaContract() As Long
    CriteriaContract = CLng(NamedValue("기준_도급", 1))
End Function

Public Function CriteriaPatrol() As Long
    CriteriaPatrol = CLng(NamedValue("기준_순회", 7))
End Function

Public Function CriteriaJoint() As Long
    CriteriaJoint = CLng(NamedValue("기준_합동", 90))
End Function

'--- 경로 --------------------------------------------------------------------

' 템플릿 폴더. 상대경로(.\form\templates\)면 통합문서 위치를 기준으로 푼다.
'
' 서식 한 장마다 불리므로(순회 12장이면 12번) 한 번 푼 것은 들고 있는다.
' ResetLookupCache 가 버튼마다 비우므로 설정을 고쳐도 곧 반영된다.
Public Function TemplateFolder() As String
    If CachedTemplateFolder <> "" Then
        TemplateFolder = CachedTemplateFolder
        Exit Function
    End If

    Dim p As String
    p = Trim$(CStr(NamedValue("SET_템플릿폴더", ".\form\templates\")))
    If Left$(p, 1) = "." Then
        p = ThisWorkbook.Path & Mid$(p, 2)
    ElseIf InStr(p, ":") = 0 And Left$(p, 2) <> "\\" Then
        p = ThisWorkbook.Path & "\" & p
    End If
    If Right$(p, 1) <> "\" Then p = p & "\"
    CachedTemplateFolder = p
    TemplateFolder = p
End Function

' 렌더링 임시 폴더. 없으면 만든다.
Public Function TempFolder() As String
    If CachedTempFolder <> "" Then
        TempFolder = CachedTempFolder
        Exit Function
    End If

    Dim p As String
    p = Trim$(CStr(NamedValue("SET_임시폴더", "")))
    If p = "" Then
        ' 통합문서 옆의 '출력' 폴더에 만든다. %TEMP% 를 쓰던 것을 옮겼다 ―
        ' 한글의 '다른 이름으로 저장' 기본 폴더가 마지막으로 연 폴더라서,
        ' 결과물이 템플릿 폴더 근처에 있으면 원본을 덮어쓸 위험이 있었다.
        Dim base As String
        base = ThisWorkbook.Path
        If Trim$(base) = "" Then base = Environ$("TEMP")
        If Right$(base, 1) = "\" Then base = Left$(base, Len(base) - 1)
        p = base & "\출력"
    End If
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)

    ' 폴더가 이미 있는데 MkDir 를 걸면 오류 75 가 난다. Dir$(..., vbDirectory) 로
    ' 확인하던 것을 FileSystemObject 로 바꿨다 ― 한셀에서 폴더를 못 찾는 일이 있어,
    ' 있는 폴더에 MkDir 가 걸리면 그 오류가 RenderTemplate 까지 튀어 올라가
    ' '서식 처리 중 오류' 로 둔갑했다.
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then
        If Dir$(p, vbDirectory) = "" Then MkDir p
    ElseIf Not fso.FolderExists(p) Then
        fso.CreateFolder p
    End If
    On Error GoTo 0

    CachedTempFolder = p & "\"
    TempFolder = CachedTempFolder
End Function

'--- 작업유형 키워드 사전 -----------------------------------------------------

' 계약명에서 작업유형을 추론한다. 반환: Boolean(0 To 3) = 고소/화기/밀폐/전기
Public Function DetectWorkTypes(ByVal contractName As String) As Variant
    Dim hit(0 To 3) As Boolean
    Dim nm As String
    nm = Replace(Trim$(contractName), " ", "")
    If nm = "" Then
        DetectWorkTypes = hit
        Exit Function
    End If

    Dim rng As Range: Set rng = NamedRange("SET_키워드")
    Dim rows As Variant
    If rng Is Nothing Then
        rows = DefaultWorkKeywords()
    Else
        rows = RangeToPairs(rng)
    End If

    Dim i As Long, idx As Long
    For i = LBound(rows, 1) To UBound(rows, 1)
        idx = WorkTypeIndex(CStr(rows(i, 0)))
        If idx >= 0 Then
            If MatchesAnyKeyword(nm, CStr(rows(i, 1))) Then hit(idx) = True
        End If
    Next i

    DetectWorkTypes = hit
End Function

Private Function WorkTypeIndex(ByVal name As String) As Long
    Select Case Replace(Trim$(name), " ", "")
        Case "고소작업", "고소":            WorkTypeIndex = 0
        Case "화기작업", "화기":            WorkTypeIndex = 1
        Case "밀폐공간작업", "밀폐공간", "밀폐": WorkTypeIndex = 2
        Case "전기작업", "전기":            WorkTypeIndex = 3
        Case Else:                          WorkTypeIndex = -1
    End Select
End Function

Private Function MatchesAnyKeyword(ByVal packedName As String, ByVal csv As String) As Boolean
    Dim parts() As String, i As Long, kw As String
    parts = Split(Replace(csv, "，", ","), ",")
    For i = LBound(parts) To UBound(parts)
        kw = Replace(Trim$(parts(i)), " ", "")
        If kw <> "" Then
            If InStr(1, packedName, kw, vbTextCompare) > 0 Then
                MatchesAnyKeyword = True
                Exit Function
            End If
        End If
    Next i
End Function

' 2열 범위를 (n,2) 배열로. 0-based 두 번째 차원.
Private Function RangeToPairs(ByVal rng As Range) As Variant
    Dim n As Long: n = rng.Rows.Count
    Dim out() As Variant
    ReDim out(1 To n, 0 To 1)
    Dim i As Long
    For i = 1 To n
        out(i, 0) = CStr(rng.Cells(i, 1).Value)
        out(i, 1) = CStr(rng.Cells(i, 2).Value)
    Next i
    RangeToPairs = out
End Function

' 설정 시트가 없을 때 쓰는 기본 사전
' (경기도교육청 중대재해처벌법 업무 매뉴얼 '참고3 주요 도급·용역·위탁 사업' 표 기준)
Public Function DefaultWorkKeywords() As Variant
    Dim out(1 To 4, 0 To 1) As Variant
    out(1, 0) = "고소작업"
    out(1, 1) = "외벽,지붕,방수,유리창,창호,전광판,도장,페인트,천장,옥상,비계,사다리,고소," & _
                "승강기,엘리베이터,조명,등기구,냉난방,에어컨,실외기,차광막,차광망,현수막," & _
                "CCTV,태양광,간판,낙하물,크레인,고가,미화"
    out(2, 0) = "화기작업"
    out(2, 1) = "용접,용단,절단,배관,보일러,난방,급식,조리,주방,후드,철거,해체,산소," & _
                "그라인더,화기,가스,열처리,방수공사"
    out(3, 0) = "밀폐공간작업"
    out(3, 1) = "물탱크,저수조,정화조,맨홀,오수,하수,배수,집수정,탱크,지하,피트,덕트," & _
                "밀폐,준설,수조,관로"
    out(4, 0) = "전기작업"
    out(4, 1) = "전기,수배전,변압기,배전반,분전반,통신,전산,네트워크,소방,승강기,태양광," & _
                "충전,누전,정전,케이블,자동제어"
    DefaultWorkKeywords = out
End Function

'--- 안전보건관리계획서 템플릿 선택 -------------------------------------------

' 계약 성격에 따라 계획서 템플릿 파일명을 고른다.
' 우선순위: 사용자가 T열에서 직접 고른 유형 > 90일↑ 또는 위험작업 2종↑ = 상세 > 공사/용역 키워드
Public Function ResolvePlanTemplate(ByVal inputRow As Long) As String
    Dim ws As Worksheet: Set ws = WsInput()
    Dim choice As String
    choice = Trim$(CStr(ws.Cells(inputRow, CC_계획서유형).Value))

    Select Case choice
        Case "상세": ResolvePlanTemplate = "01_계획서_상세.hwpx": Exit Function
        Case "공사": ResolvePlanTemplate = "01_계획서_공사.hwpx": Exit Function
        Case "용역": ResolvePlanTemplate = "01_계획서_용역.hwpx": Exit Function
    End Select

    ' 자동 판정
    Dim days As Long: days = ContractDays(inputRow)
    Dim types As String: types = ActiveWorkTypes(inputRow)
    Dim typeCount As Long
    If types <> "" Then typeCount = UBound(Split(types, "|")) + 1

    If days >= 90 Or typeCount >= 2 Then
        ResolvePlanTemplate = "01_계획서_상세.hwpx"
        Exit Function
    End If

    Dim nm As String
    nm = Replace(CStr(ws.Cells(inputRow, CC_계약명).Value), " ", "")
    If MatchesAnyKeyword(nm, PlanKeywords("공사")) Then
        ResolvePlanTemplate = "01_계획서_공사.hwpx"
    ElseIf MatchesAnyKeyword(nm, PlanKeywords("용역")) Then
        ResolvePlanTemplate = "01_계획서_용역.hwpx"
    Else
        ResolvePlanTemplate = "01_계획서_상세.hwpx"
    End If
End Function

Private Function PlanKeywords(ByVal kind As String) As String
    Dim rng As Range: Set rng = NamedRange("SET_계획서")
    If Not rng Is Nothing Then
        Dim i As Long
        For i = 1 To rng.Rows.Count
            If Trim$(CStr(rng.Cells(i, 1).Value)) = kind Then
                PlanKeywords = CStr(rng.Cells(i, 2).Value)
                Exit Function
            End If
        Next i
    End If
    If kind = "공사" Then
        PlanKeywords = "공사,설치,교체,보수,수선,철거,해체,시공,증축,개축,도장,방수,포장"
    Else
        PlanKeywords = "용역,위탁,청소,미화,경비,급식,조리,유지관리,유지보수,점검,관리,운영"
    End If
End Function

'--- 점검요일 기본값 ----------------------------------------------------------

Public Function DefaultInspectWeekday() As String
    DefaultInspectWeekday = CStr(NamedValue("SET_기본점검요일", "수요일"))
End Function

'--- 추천 안내창 ---------------------------------------------------------------

' 계약명을 넣었을 때 '이렇게 추천했습니다' 창을 띄울지. 설정 시트에서 끌 수 있다.
' 이름이 없으면 켜 둔 것으로 본다 ― 처음 쓰는 사람에게 알려 주는 것이 목적이다.
Public Function ShowDetectNoticeEnabled() As Boolean
    Dim v As String
    v = Trim$(CStr(NamedValue("SET_추천안내", "표시")))
    ShowDetectNoticeEnabled = (v <> "숨김")
End Function

' "월요일" → vbMonday(2). 알 수 없으면 vbWednesday(4).
Public Function WeekdayNumber(ByVal name As String) As Long
    Select Case Trim$(name)
        Case "일요일": WeekdayNumber = 1
        Case "월요일": WeekdayNumber = 2
        Case "화요일": WeekdayNumber = 3
        Case "수요일": WeekdayNumber = 4
        Case "목요일": WeekdayNumber = 5
        Case "금요일": WeekdayNumber = 6
        Case "토요일": WeekdayNumber = 7
        Case Else:     WeekdayNumber = 4
    End Select
End Function

Public Function WeekdayName2(ByVal wd As Long) As String
    Select Case wd
        Case 1: WeekdayName2 = "일요일"
        Case 2: WeekdayName2 = "월요일"
        Case 3: WeekdayName2 = "화요일"
        Case 4: WeekdayName2 = "수요일"
        Case 5: WeekdayName2 = "목요일"
        Case 6: WeekdayName2 = "금요일"
        Case 7: WeekdayName2 = "토요일"
    End Select
End Function

'==============================================================================
' 공휴일 ― 점검일이 공휴일이면 다음 평일로 미룬다
'
' 설날·추석·부처님오신날은 음력이라 계산으로 낼 수 없다. 그래서 날짜를 적어 둔다.
' 대체공휴일·선거일·근로자의날(5/1)·제헌절(7/17)까지 관공서가 쉬는 날은 모두 넣었다.
'
' 표는 2030년까지다. 그 뒤 날짜는 공휴일 판정을 하지 않고 그냥 지나간다
' (점검일이 하루 어긋날 뿐 출력은 된다). 해가 바뀌면 여기에 한 줄 더 붙이면 된다.
'==============================================================================

Private Function HolidayList() As String
    Dim h As String
    h = h & " 20260101 20260216 20260217 20260218 20260301 20260302 20260501 20260505 20260524 20260525 20260603 20260606 20260717 20260815 20260817 20260924 20260925 20260926 20261003 20261005 20261009 20261225"    ' 2026
    h = h & " 20270101 20270206 20270207 20270208 20270209 20270301 20270501 20270503 20270505 20270513 20270606 20270717 20270719 20270815 20270816 20270914 20270915 20270916 20271003 20271004 20271009 20271011 20271225 20271227"    ' 2027
    h = h & " 20280101 20280126 20280127 20280128 20280301 20280412 20280501 20280502 20280505 20280606 20280717 20280815 20281002 20281003 20281004 20281005 20281009 20281225"    ' 2028
    h = h & " 20290101 20290212 20290213 20290214 20290301 20290501 20290505 20290507 20290520 20290521 20290606 20290717 20290815 20290921 20290922 20290923 20290924 20291003 20291009 20291225"    ' 2029
    h = h & " 20300101 20300202 20300203 20300204 20300205 20300301 20300403 20300501 20300505 20300506 20300509 20300606 20300612 20300717 20300815 20300911 20300912 20300913 20301003 20301009 20301225"    ' 2030
    HolidayList = h
End Function

Public Function IsHoliday(ByVal d As Date) As Boolean
    If HolidayCache Is Nothing Then
        Set HolidayCache = CreateObject("Scripting.Dictionary")
        Dim p() As String, i As Long
        p = Split(Trim$(HolidayList()), " ")
        For i = LBound(p) To UBound(p)
            If p(i) <> "" Then HolidayCache(p(i)) = True
        Next i
    End If
    IsHoliday = HolidayCache.Exists(Format$(d, "yyyymmdd"))
End Function

' 주말이거나 공휴일이면 다음 평일로 미룬다. 연휴가 길어도 계속 민다.
' (표에 없는 해는 주말만 피한다)
Public Function NextWorkday(ByVal d As Date) As Date
    Dim i As Long
    Do While Weekday(d) = vbSunday Or Weekday(d) = vbSaturday Or IsHoliday(d)
        d = d + 1
        i = i + 1
        If i > 30 Then Exit Do          ' 무슨 일이 있어도 멈춘다
    Loop
    NextWorkday = d
End Function

'--- 기관구분별 직위 기본값 ---------------------------------------------------
'
' 결재선의 부서장과 기관대표자는 다르다. 학교는 둘 다 교장이지만
' 교육지원청은 부서장이 과장, 기관대표자는 교육장이다.
' idx: 0=담당자 1=관리자 2=부서장 3=기관대표자
Public Function OrgPositionDefaults(ByVal orgKind As String) As Variant
    If InStr(orgKind, "지원청") > 0 Or InStr(orgKind, "교육청") > 0 Then
        OrgPositionDefaults = Array("주무관", "팀장", "과장", "교육장")
    Else
        OrgPositionDefaults = Array("주무관", "행정실장", "교장", "교장")
    End If
End Function

'==============================================================================
' 계약유형 사전 ― 계약명으로 그 계약에 맞는 문구 묶음을 찾는다
'
' 시트 '계약유형사전' 의 한 줄이 계약 한 종류다.
'   A 유형명 / B 계약명 키워드 / C 주요 작업내용 / D 위험요인 / E 안전대책
'   F 협의체 안건 / G 점검 착안사항 / H 기본 작업유형(고소|화기|밀폐|전기)
' 사용자가 시트에서 문구를 직접 고칠 수 있도록 숨기지 않는다.
' (열 번호 상수 DICT_* 는 모듈 맨 위 선언부에 있다)
'==============================================================================

Private Function DictRange() As Range
    Set DictRange = NamedRange("SET_계약유형사전")
End Function

'--- 사전을 통째로 떠 두기 ----------------------------------------------------
'
' ContractTypeField 는 BuildTagMap 한 번에 스무 번 넘게 불린다. 그때마다 사전을
' 셀 하나씩 물으면 계약 한 건에 수천 번 COM 을 오간다. 배열로 한 번에 떠서
' 그 위에서 찾는다.
'
' 사용자가 사전 시트를 고치고 곧바로 출력할 수 있으므로 오래 들고 있으면 안 된다.
' ResetLookupCache 가 버튼을 누를 때마다 비운다.
' (DictCache · DictIndex · CachedTemplateFolder · CachedTempFolder 는 모듈 맨 위
'  선언부에 있다 ― VBA 는 모듈 수준 변수를 첫 프로시저보다 위에서만 받는다)

Private Sub LoadDictCache()
    If Not IsEmpty(DictCache) Then Exit Sub

    Dim rng As Range: Set rng = DictRange()
    If rng Is Nothing Then Exit Sub
    If rng.Cells.Count = 1 Then Exit Sub      ' 한 칸짜리면 배열이 안 나온다. 그냥 둔다.

    DictCache = rng.Value

    Set DictIndex = CreateObject("Scripting.Dictionary")
    Dim i As Long, nm As String
    For i = 1 To UBound(DictCache, 1)
        nm = Trim$(CStr(DictCache(i, DICT_유형명) & ""))
        ' 같은 유형명이 두 줄이면 위의 것을 쓴다 ― 예전 루프도 먼저 걸린 줄에서 멈췄다.
        If nm <> "" And Not DictIndex.Exists(nm) Then DictIndex(nm) = i
    Next i
End Sub

' 시트에서 떠 온 것을 모두 버린다. modHwpExport 의 ResetExportCache 가 부른다.
Public Sub ResetLookupCache()
    DictCache = Empty
    Set DictIndex = Nothing
    CachedTemplateFolder = ""
    CachedTempFolder = ""
End Sub

' 계약명에 맞는 사전 줄을 찾는다. 못 찾으면 Nothing.
'
' 키워드가 여러 줄에 걸릴 수 있다(예: '소방시설 정기점검' 은 소방·점검 둘 다 걸린다).
' 가장 긴 키워드가 걸린 줄을 고른다. 긴 키워드일수록 구체적이기 때문이다.
Public Function FindContractType(ByVal contractName As String) As Range
    Dim rng As Range: Set rng = DictRange()
    If rng Is Nothing Then Exit Function

    Dim nm As String
    nm = Replace(Trim$(contractName), " ", "")
    If nm = "" Then Exit Function

    LoadDictCache

    ' 줄 번호만 메모리에서 고르고, 마지막에 한 번만 Range 로 바꾼다.
    Dim bestRow As Long, bestLen As Long
    Dim i As Long, hitLen As Long, rowCount As Long

    If IsEmpty(DictCache) Then
        rowCount = rng.Rows.Count                 ' 사전을 못 떴으면 예전처럼 셀로 읽는다
    Else
        rowCount = UBound(DictCache, 1)
    End If

    For i = 1 To rowCount
        If Trim$(CStr(DictField(rng, i, DICT_유형명))) <> "" Then
            hitLen = LongestKeywordHit(nm, CStr(DictField(rng, i, DICT_키워드)))
            If hitLen > bestLen Then
                bestLen = hitLen
                bestRow = i
            End If
        End If
    Next i

    If bestRow > 0 Then Set FindContractType = rng.Rows(bestRow)
End Function

' 사전의 한 칸. 떠 둔 배열이 있으면 거기서, 없으면 시트에서 읽는다.
'
' 이름 정의(A:T)보다 오른쪽 칸을 물으면 배열에는 없다. 그럴 때는 시트에서 읽는다 ―
' Range.Cells 는 범위 밖도 그대로 짚어 주므로, 셀을 하나씩 읽던 예전과 값이 같다.
Private Function DictField(ByVal rng As Range, ByVal row As Long, ByVal col As Long) As Variant
    If IsEmpty(DictCache) Then
        DictField = rng.Cells(row, col).Value
    ElseIf col > UBound(DictCache, 2) Then
        DictField = rng.Cells(row, col).Value
    Else
        DictField = DictCache(row, col)
    End If
End Function

' 계약명에 걸린 키워드 중 가장 긴 것의 길이. 하나도 안 걸리면 0.
'
' 길이를 지역 변수(best)에 모은다. 함수 이름을 값처럼 읽는 것(If Len(kw) > LongestKeywordHit)
' 은 엑셀 VBA 에서는 되지만 **한셀에서는 컴파일이 깨진다** ― 매개변수가 있는 함수라
' 이름만 쓰면 '생략될 수 없는 매개변수입니다' 로 본다. 함수 이름은 오직 결과를
' 대입할 때만 쓴다.
Private Function LongestKeywordHit(ByVal packedName As String, ByVal csv As String) As Long
    Dim parts() As String, i As Long, kw As String, best As Long
    parts = Split(Replace(csv, "，", ","), ",")
    For i = LBound(parts) To UBound(parts)
        kw = Replace(Trim$(parts(i)), " ", "")
        If kw <> "" Then
            If InStr(1, packedName, kw, vbTextCompare) > 0 Then
                If Len(kw) > best Then best = Len(kw)
            End If
        End If
    Next i
    LongestKeywordHit = best
End Function

' 계약명 → 유형명 ("" 이면 못 찾음)
Public Function ContractTypeName(ByVal contractName As String) As String
    Dim r As Range: Set r = FindContractType(contractName)
    If Not r Is Nothing Then ContractTypeName = Trim$(CStr(r.Cells(1, DICT_유형명).Value))
End Function

' 유형명 → 그 줄의 한 칸 값
Public Function ContractTypeField(ByVal typeName As String, ByVal col As Long) As String
    Dim rng As Range: Set rng = DictRange()
    If rng Is Nothing Or Trim$(typeName) = "" Then Exit Function

    LoadDictCache

    ' 떠 둔 사전이 있으면 유형명 → 줄 번호를 바로 찾는다(훑지 않는다).
    If Not DictIndex Is Nothing Then
        If DictIndex.Exists(Trim$(typeName)) Then
            ContractTypeField = Trim$(CStr(DictField(rng, DictIndex(Trim$(typeName)), col) & ""))
        End If
        Exit Function
    End If

    Dim i As Long
    For i = 1 To rng.Rows.Count
        If Trim$(CStr(rng.Cells(i, DICT_유형명).Value)) = Trim$(typeName) Then
            ContractTypeField = Trim$(CStr(rng.Cells(i, col).Value))
            Exit Function
        End If
    Next i
End Function

' 계약유형 사전이 정해 둔 기본 작업유형. 반환: Boolean(0 To 3), 사전에 없으면 모두 False
Public Function DictWorkTypes(ByVal typeName As String) As Variant
    Dim hit(0 To 3) As Boolean
    Dim spec As String
    spec = ContractTypeField(typeName, DICT_작업유형)
    If spec <> "" Then
        Dim parts() As String, i As Long, idx As Long
        parts = Split(Replace(Replace(spec, ",", "|"), " ", ""), "|")
        For i = LBound(parts) To UBound(parts)
            idx = WorkTypeIndex(parts(i))
            If idx >= 0 Then hit(idx) = True
        Next i
    End If
    DictWorkTypes = hit
End Function
