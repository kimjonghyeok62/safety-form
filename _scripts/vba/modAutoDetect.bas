Attribute VB_Name = "modAutoDetect"
Option Explicit

'==============================================================================
' 계약명으로 작업유형(고소·화기·밀폐·전기)을 자동 판정한다.
'
' 사용자가 손으로 고친 칸은 절대 덮어쓰지 않는다. 구분은 셀 배경색으로 한다.
'   - 배경색 = AUTO_FILL_COLOR  → 자동으로 채운 칸. 계약명이 바뀌면 다시 계산한다.
'   - 배경색 없음(또는 다른 색) → 사람이 손댄 칸. 건드리지 않는다.
' 이 규칙은 눈으로도 바로 보이므로 사용자가 "이건 자동, 이건 내가 정한 것"을
' 따로 기억할 필요가 없다.
'==============================================================================

' 한 계약 행의 작업유형을 다시 판정한다.
'
' 판정 순서: ① 계약유형 사전에 걸리면 그 줄이 정해 둔 작업유형을 쓴다
'            ② 사전에 없으면 설정 시트의 낱말 사전으로 추론한다
' 사전 쪽이 사람이 정리해 둔 값이라 더 정확하다.
Public Sub RefreshWorkTypes(ByVal inputRow As Long)
    If inputRow < IN_FIRST_ROW Or inputRow > IN_LAST_ROW Then Exit Sub

    Dim ws As Worksheet: Set ws = WsInput()
    Dim nm As String
    nm = Trim$(CStr(ws.Cells(inputRow, CC_계약명).Value))

    Dim hit As Variant
    hit = DetectWorkTypes(nm)

    Dim tname As String
    tname = Trim$(CStr(ws.Cells(inputRow, CC_계약유형).Value))
    If tname <> "" Then
        Dim dictHit As Variant
        dictHit = DictWorkTypes(tname)
        Dim k As Long, anyDict As Boolean
        For k = 0 To 3
            If dictHit(k) Then anyDict = True
        Next k
        If anyDict Then hit = dictHit
    End If

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False        ' 자동 채움이 Change 이벤트를 다시 부르지 않도록

    Dim c As Long, cel As Range
    For c = WORK_FIRST_COL To WORK_LAST_COL
        Set cel = ws.Cells(inputRow, c)
        If IsAutoCell(cel) Then
            If nm = "" Then
                cel.Value = ""
                cel.Interior.ColorIndex = xlNone
            ElseIf hit(c - WORK_FIRST_COL) Then
                cel.Value = "○"
                cel.Interior.Color = AUTO_FILL_COLOR
            Else
                ' 빈칸이 아니라 × 를 찍는다. '판정한 결과 아님' 과 '아직 판정 전' 을
                ' 눈으로 구분할 수 있어야 한다.
                cel.Value = "×"
                cel.Interior.Color = AUTO_FILL_COLOR
            End If
        End If
    Next c

    Application.EnableEvents = prevEvents
End Sub

' 자동 판정 대상 칸인가?
'   노란색   = 자동으로 채운 칸        → 다시 계산한다
'   회색     = 사용자가 직접 정한 칸    → 손대지 않는다
'   색 없음  = 아직 아무도 손대지 않은 칸 → 자동 대상
' '색 없음'과 '사용자가 비운 칸'을 색으로 구분하지 않으면, 사용자가 지운 ○ 가
' 계약명을 고칠 때마다 되살아난다.
Private Function IsAutoCell(ByVal cel As Range) As Boolean
    If cel.Interior.ColorIndex = xlNone Then
        IsAutoCell = True
    Else
        IsAutoCell = (cel.Interior.Color = AUTO_FILL_COLOR)
    End If
End Function

' 사용자가 작업유형 칸을 직접 고쳤다 → 회색으로 칠해 '수동'으로 못박는다.
Public Sub MarkManual(ByVal cel As Range)
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    cel.Interior.Color = MANUAL_FILL_COLOR
    Application.EnableEvents = prevEvents
End Sub

' 자동 판정으로 되돌린다(선택 영역의 '수동' 표시를 지우고 다시 계산).
Public Sub ResetToAuto()
    Dim ws As Worksheet: Set ws = WsInput()
    Dim sel As Range
    On Error Resume Next
    Set sel = Application.Intersect(Selection, _
              ws.Range(ws.Cells(IN_FIRST_ROW, WORK_FIRST_COL), _
                       ws.Cells(IN_LAST_ROW, WORK_LAST_COL)))
    On Error GoTo 0

    If sel Is Nothing Then
        MsgBox "입력정보 시트에서 작업유형(고소·화기·밀폐·전기) 칸을 먼저 선택하세요.", _
               vbInformation, "자동 판정으로 되돌리기"
        Exit Sub
    End If

    Application.EnableEvents = False
    sel.Interior.ColorIndex = xlNone
    Application.EnableEvents = True

    Dim r As Long
    For r = IN_FIRST_ROW To IN_LAST_ROW
        RefreshWorkTypes r
    Next r
End Sub

' 계약명이 처음 입력될 때 점검요일 기본값을 채워 준다(이미 값이 있으면 두고 본다).
Public Sub EnsureInspectWeekday(ByVal inputRow As Long)
    Dim ws As Worksheet: Set ws = WsInput()
    If Trim$(CStr(ws.Cells(inputRow, CC_계약명).Value)) = "" Then Exit Sub
    If Trim$(CStr(ws.Cells(inputRow, CC_점검요일).Value)) <> "" Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    ws.Cells(inputRow, CC_점검요일).Value = DefaultInspectWeekday()
    Application.EnableEvents = prevEvents
End Sub

' 계획서 유형 칸이 비어 있으면 '자동'으로 채운다.
Public Sub EnsurePlanType(ByVal inputRow As Long)
    Dim ws As Worksheet: Set ws = WsInput()
    If Trim$(CStr(ws.Cells(inputRow, CC_계약명).Value)) = "" Then Exit Sub
    If Trim$(CStr(ws.Cells(inputRow, CC_계획서유형).Value)) <> "" Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    ws.Cells(inputRow, CC_계획서유형).Value = "자동"
    Application.EnableEvents = prevEvents
End Sub

' 계약명으로 계약유형 사전을 뒤져 S열을 채운다. 작업유형과 같은 3색 규칙을 쓴다.
Public Sub RefreshContractType(ByVal inputRow As Long)
    If inputRow < IN_FIRST_ROW Or inputRow > IN_LAST_ROW Then Exit Sub

    Dim ws As Worksheet: Set ws = WsInput()
    Dim cel As Range: Set cel = ws.Cells(inputRow, CC_계약유형)
    If Not IsAutoCell(cel) Then Exit Sub        ' 사용자가 직접 정한 값은 두고 본다

    Dim nm As String
    nm = Trim$(CStr(ws.Cells(inputRow, CC_계약명).Value))

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    If nm = "" Then
        cel.Value = ""
        cel.Interior.ColorIndex = xlNone
    Else
        cel.Value = ContractTypeName(nm)
        cel.Interior.Color = AUTO_FILL_COLOR
    End If

    Application.EnableEvents = prevEvents
End Sub

'==============================================================================
' 추천 안내창
'
' 계약명을 넣으면 네 가지(계약유형·작업유형·계획서유형·점검요일)가 자동으로
' 정해지는데, 그 사실이 화면에 드러나지 않는다. 노란 칸이 '자동으로 채운 값'
' 이라는 규칙도 처음 쓰는 사람은 모른다. 그래서 한 번 알려 준다.
'
' 확인을 눌러야 닫히는 창(MsgBox)으로 만들면 계약 100건을 넣는 동안 100번을
' 눌러야 한다. 그래서 스스로 닫히는 안내창(frmNotice)을 쓴다.
'==============================================================================

Public Sub ShowDetectNotice(ByVal inputRow As Long)
    If SilentMode Then Exit Sub
    If Not ShowDetectNoticeEnabled() Then Exit Sub

    Dim ws As Worksheet: Set ws = WsInput()
    Dim nm As String
    nm = Trim$(CStr(ws.Cells(inputRow, CC_계약명).Value))
    If nm = "" Then Exit Sub

    Dim tname As String, types As String
    tname = Trim$(CStr(ws.Cells(inputRow, CC_계약유형).Value))
    types = Replace(ActiveWorkTypes(inputRow), "|", ", ")

    Dim body As String
    body = "계약명: " & nm & vbLf & vbLf & _
           "   계약유형     " & IIf(tname = "", "(사전에서 찾지 못했습니다)", tname) & vbLf & _
           "   작업유형     " & IIf(types = "", "(해당 없음)", types) & vbLf & _
           "   계획서유형   " & PlanKind(inputRow) & vbLf & _
           "   점검요일     " & CStr(ws.Cells(inputRow, CC_점검요일).Value) & vbLf & vbLf & _
           "이렇게 추천해 두었습니다. 그대로 두셔도 되고 바꾸셔도 됩니다." & vbLf & _
           "연노랑 칸이 자동으로 채운 값입니다. 그 위에 직접 쓰시면 연회색으로" & vbLf & _
           "바뀌고, 그다음부터는 자동 판정이 그 칸을 건드리지 않습니다." & vbLf & vbLf & _
           "(이 창은 잠시 뒤 저절로 닫힙니다. 설정 시트에서 끌 수 있습니다)"

    ShowNotice "자동 추천", body

    ' 안내창이 초점을 가져가면 이어서 치던 글자가 어디로도 가지 않는다. 되돌린다.
    On Error Resume Next
    AppActivate Application.Caption
    Application.OnTime Now + TimeSerial(0, 0, 6), "HideNotice"
    On Error GoTo 0
End Sub

' 기관구분(학교/교육지원청)을 고르면 기관정보 두 줄을 그 기관 것으로 갈아 끼운다.
'
'   직위 4칸   ― 늘 기본값으로 덮어쓴다. 기관구분이 정하는 값이라 되살릴 이유가 없다.
'                잠그지는 않는다 ― 그냥 타이핑하면 덮어써진다.
'                (예전에는 '자동으로 넣었음' 표시로 연회색을 칠했는데, 지금은
'                 조건부서식이 '빈칸=노랑 / 채움=회색' 을 늘 덮어써서 보이지 않는다.
'                 보이지 않는 표시를 남겨 두면 코드만 헷갈리므로 없앴다)
'   나머지 전부 ― 기관명·고유번호·주소·대표전화·팩스번호와 결재선 성명·전화.
'                기관이 바뀌면 학교 이름부터 사람까지 다 달라지므로 통째로 비운다.
'                다만 비우기 전에 보관함에 넣어 두고, 그 기관으로 되돌아오면 되살린다.
'                실수로 잘못 골랐다가 되돌렸을 때 다시 치게 하면 안 되기 때문이다.
Public Sub ApplyOrgDefaults()
    Dim ws As Worksheet: Set ws = WsInput()
    Dim kind As String
    kind = Trim$(CStr(ws.Cells(IN_SCHOOL_ROW, IC_기관구분).Value))
    If kind = "" Then Exit Sub

    Dim d As Variant
    d = OrgPositionDefaults(kind)

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    ' ① 직위 4칸은 무조건 기본값으로
    SetOrgPosition ws.Cells(IN_SIGN_ROW, SC_담당자직위), CStr(d(0))
    SetOrgPosition ws.Cells(IN_SIGN_ROW, SC_관리자직위), CStr(d(1))
    SetOrgPosition ws.Cells(IN_SIGN_ROW, SC_부서장직위), CStr(d(2))
    SetOrgPosition ws.Cells(IN_SIGN_ROW, SC_대표자직위), CStr(d(3))

    ' ② 기관구분이 실제로 바뀌었을 때만 성명·전화를 갈아 끼운다.
    '    (같은 값을 다시 고른 것이라면 애써 적은 이름을 지울 이유가 없다)
    Dim prev As String
    prev = StashState()
    If prev = "" Then
        ' 전에 무엇이었는지 모르면 지금 적힌 성명·전화가 어느 기관 것인지도 모른다.
        ' 함부로 비우지 말고 지금 기관구분 것으로 보고 적어만 둔다.
        SetStashState kind
        StashSave kind
    ElseIf prev <> kind Then
        StashSave prev
        StashLoad kind                      ' 보관해 둔 것이 없으면 비운다
        SetStashState kind
    End If

    ' 학교는 부서장과 기관대표자가 같은 사람(교장)이라 성명을 따라 채운다.
    ' 교육지원청은 과장과 교육장이 다르므로 건드리지 않는다.
    If CStr(d(2)) = CStr(d(3)) Then
        FillIfBlank ws.Cells(IN_SIGN_ROW, SC_대표자성명), _
                    CStr(ws.Cells(IN_SIGN_ROW, SC_부서장성명).Value)
    End If

    Application.EnableEvents = prevEvents
End Sub

Private Sub FillIfBlank(ByVal cel As Range, ByVal v As String)
    If Trim$(CStr(cel.Value)) = "" Then cel.Value = v
End Sub

' 직위 한 칸을 기본값으로 덮어쓴다. 배경은 조건부서식이 맡는다.
Private Sub SetOrgPosition(ByVal cel As Range, ByVal v As String)
    cel.Value = v
End Sub

'------------------------------------------------------------------------------
' 결재선 보관함 ― ContentLibrary 시트 AA열 (숨김)
'------------------------------------------------------------------------------

' 보관·복원 대상 칸. 기관이 바뀌면 기관정보도 결재선 사람도 통째로 달라지므로
' 두 줄을 다 담는다. 기관구분 칸과 직위 4칸만 뺀다(그 둘은 기관구분이 정하는 값이다).
' 저장 순서는 이 배열 순서를 따른다.
Private Function StashSchoolCols() As Variant
    StashSchoolCols = Array(IC_기관명, IC_고유번호, IC_주소, IC_전화번호, IC_팩스번호)
End Function

Private Function StashSignCols() As Variant
    StashSignCols = Array(SC_담당자성명, SC_담당자전화, SC_관리자성명, _
                          SC_부서장성명, SC_대표자성명)
End Function

Private Function WsStash() As Worksheet
    On Error Resume Next
    Set WsStash = ThisWorkbook.Worksheets(SH_LIB)
    On Error GoTo 0
End Function

' 지금 적용 중인 기관구분을 읽는다.
'
' 예전에는 이 함수 하나가 Optional 인수로 읽기와 쓰기를 겸했다. 그런데 쓸 때는
' 반환값이 필요 없어 'StashState kind' 처럼 문장으로 불렀는데, 그것은 Function 을
' 문장처럼 부르는 꼴이라 한셀이 여는 괄호를 찾다가 컴파일을 멈춘다(엑셀은 받아 준다).
' 읽기는 Function, 쓰기는 Sub 로 갈라 둔다.
Private Function StashState() As String
    Dim ws As Worksheet: Set ws = WsStash()
    If ws Is Nothing Then Exit Function
    StashState = Trim$(CStr(ws.Cells(STASH_STATE_ROW, STASH_COL).Value))
End Function

' 지금 적용 중인 기관구분을 적어 둔다.
Private Sub SetStashState(ByVal newKind As String)
    If newKind = "" Then Exit Sub
    Dim ws As Worksheet: Set ws = WsStash()
    If ws Is Nothing Then Exit Sub
    ws.Cells(STASH_STATE_ROW, STASH_COL).Value = newKind
End Sub

' 그 기관구분이 쓰는 보관 줄. 없으면 빈 줄을 하나 잡아 준다. 자리가 없으면 0.
Private Function StashRow(ByVal kind As String) As Long
    Dim ws As Worksheet: Set ws = WsStash()
    If ws Is Nothing Then Exit Function

    Dim r As Long, v As String
    For r = STASH_FIRST_ROW To STASH_LAST_ROW
        v = Trim$(CStr(ws.Cells(r, STASH_COL).Value))
        If v = kind Then
            StashRow = r
            Exit Function
        ElseIf v = "" Then
            ws.Cells(r, STASH_COL).Value = kind
            StashRow = r
            Exit Function
        End If
    Next r
End Function

' 지금 기관정보·결재선 값을 그 기관구분 이름으로 넣어 둔다.
Private Sub StashSave(ByVal kind As String)
    Dim ws As Worksheet: Set ws = WsStash()
    If ws Is Nothing Then Exit Sub
    Dim r As Long: r = StashRow(kind)
    If r = 0 Then Exit Sub

    Dim wi As Worksheet: Set wi = WsInput()
    Dim sc As Variant: sc = StashSchoolCols()
    Dim gc As Variant: gc = StashSignCols()
    Dim i As Long, n As Long

    For i = LBound(sc) To UBound(sc)
        ws.Cells(r, STASH_COL + 1 + n).Value = wi.Cells(IN_SCHOOL_ROW, sc(i)).Value
        n = n + 1
    Next i
    For i = LBound(gc) To UBound(gc)
        ws.Cells(r, STASH_COL + 1 + n).Value = wi.Cells(IN_SIGN_ROW, gc(i)).Value
        n = n + 1
    Next i
End Sub

' 보관해 둔 기관정보·결재선을 되살린다. 보관해 둔 것이 없으면 칸을 비운다.
Private Sub StashLoad(ByVal kind As String)
    Dim ws As Worksheet: Set ws = WsStash()
    Dim wi As Worksheet: Set wi = WsInput()
    Dim sc As Variant: sc = StashSchoolCols()
    Dim gc As Variant: gc = StashSignCols()
    Dim i As Long, n As Long, r As Long

    If Not ws Is Nothing Then r = StashRow(kind)

    For i = LBound(sc) To UBound(sc)
        FillFromStash wi.Cells(IN_SCHOOL_ROW, sc(i)), ws, r, n
        n = n + 1
    Next i
    For i = LBound(gc) To UBound(gc)
        FillFromStash wi.Cells(IN_SIGN_ROW, gc(i)), ws, r, n
        n = n + 1
    Next i
End Sub

Private Sub FillFromStash(ByVal cel As Range, ByVal ws As Worksheet, _
                          ByVal r As Long, ByVal n As Long)
    If r = 0 Or ws Is Nothing Then
        cel.ClearContents
    Else
        cel.Value = ws.Cells(r, STASH_COL + 1 + n).Value
    End If
End Sub

' 전체 계약 행을 다시 판정한다(설정 시트의 키워드를 고친 뒤 쓰는 버튼).
Public Sub RefreshAllWorkTypes()
    RefreshAllWorkTypesSilent
    MsgBox "계약명을 바탕으로 계약유형과 작업유형을 다시 판정했습니다." & vbLf & vbLf & _
           "노란 칸 = 자동으로 채운 값" & vbLf & _
           "회색 칸 = 직접 정하신 값 (자동 판정이 건드리지 않습니다)", _
           vbInformation, "자동 판정 완료"
End Sub

' 알림창 없이 판정만 한다(스크립트에서 부를 때 쓴다).
Public Sub RefreshAllWorkTypesSilent()
    Dim r As Long
    Application.ScreenUpdating = False
    For r = IN_FIRST_ROW To IN_LAST_ROW
        RefreshContractType r        ' 작업유형 판정이 이 결과를 쓰므로 먼저 한다
        RefreshWorkTypes r
        EnsureInspectWeekday r
        EnsurePlanType r
    Next r
    Application.ScreenUpdating = True
End Sub
