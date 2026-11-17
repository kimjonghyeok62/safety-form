Attribute VB_Name = "modUI"
Option Explicit

'==============================================================================
' 출력관리·설정 시트 버튼 동작
'
' 여기 있는 Btn_* 는 **모두 시트의 도형에 걸려 있다.** 버튼이 없는 매크로는 두지
' 않는다 ― 매크로 목록에 이름만 남아 있으면 무엇이 살아 있는 기능인지 헷갈린다.
' 도형을 지웠으면 여기 프로시저도 함께 지울 것.
'==============================================================================

' 강조 중인 출력관리 행. 0이면 강조 없음.
' 모듈 수준 선언은 반드시 첫 프로시저보다 위에 있어야 한다.
Private mHighlightRow As Long

' 출력관리 E열(점검요일)에 입력정보 값을 비춘다.
'
' 수식으로 두면 편집이 안 되고, 그러면 요일을 바꾸러 입력정보까지 갔다 와야 한다.
' 그래서 값으로 써 두고, 사용자가 고치면 시트 이벤트가 입력정보로 되돌려 쓴다.
Public Sub RefreshOutputWeekdays()
    Dim ws As Worksheet: Set ws = WsOutput()
    Dim wi As Worksheet: Set wi = WsInput()

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False

    Dim r As Long, ir As Long
    For r = OUT_FIRST_ROW To OUT_LAST_ROW
        ir = InputRow(r)
        If IsContractRow(ir) Then
            ws.Cells(r, OC_점검요일).Value = wi.Cells(ir, CC_점검요일).Value
        Else
            ws.Cells(r, OC_점검요일).ClearContents
        End If
    Next r

    Application.EnableEvents = prevEvents
End Sub

' 셀 값(TRUE/FALSE)에 맞춰 체크박스 모양을 맞춘다.
'
' 한셀에서 두 가지가 걸린다. 엑셀은 둘 다 받아 주지만 여기서 프로젝트 전체의
' 컴파일이 깨지면 modUI 의 매크로를 아예 부를 수 없게 된다(도형을 눌러도
' '이 개체에 지정된 매크로를 실행할 수 없습니다' 만 나온다).
'   · CheckBoxes 는 속성이 아니라 **메서드**다. 괄호를 붙여야 한다.
'   · CheckBox 형식은 한셀에 없다. 늦은 바인딩(Object)으로 받는다.
Public Sub SyncCheckboxes()
    Dim ws As Worksheet: Set ws = WsOutput()
    Dim cb As Object
    On Error Resume Next
    For Each cb In ws.CheckBoxes()
        If Left$(cb.Name, 4) = "sel_" Then
            Dim r As Long
            r = CLng(Mid$(cb.Name, 5))
            cb.Value = IIf(ws.Cells(r, OC_선택).Value = True, xlOn, xlOff)
        End If
    Next cb
    On Error GoTo 0
End Sub

' [템플릿 폴더 열기]
Public Sub Btn_OpenTemplateFolder()
    Dim p As String
    p = TemplateFolder()
    If Dir$(p, vbDirectory) = "" Then
        MsgBox "템플릿 폴더가 없습니다:" & vbLf & p, vbExclamation, "폴더 없음"
        Exit Sub
    End If
    ' Shell 은 Sub 가 아니라 Function 이다. 반환값을 버리고 문장처럼 부르면
    ' 엑셀은 받아 주지만 한셀은 여는 괄호를 찾다가 컴파일을 멈춘다. 값을 받아 둔다.
    Dim taskId As Double
    taskId = Shell("explorer.exe """ & p & """", vbNormalFocus)
End Sub

' [계약유형사전 열기]
Public Sub Btn_GoToDict()
    Dim ws As Worksheet: Set ws = WsDict()
    If ws Is Nothing Then
        MsgBox "'계약유형사전' 시트를 찾을 수 없습니다.", vbExclamation, "계약유형사전"
        Exit Sub
    End If
    ws.Activate
End Sub

' [작업유형 다시 판정]
Public Sub Btn_RefreshDetect()
    RefreshAllWorkTypes
End Sub

'==============================================================================
' 입력정보 U열의 [▶ 출력관리] 버튼
'
' 한 번 누르면 세 가지를 한꺼번에 한다.
'   ① 출력관리로 넘어가 그 줄을 화면에 띄우고
'   ② 그 계약만 체크하고 (바로 출력 버튼을 누를 수 있는 상태)
'   ③ 그 줄을 잠깐 연두색으로 칠해 어디를 보면 되는지 알려 준다
' 예전에는 이동 → 체크 → 버튼으로 세 번을 눌러야 했다.
'==============================================================================

Public Sub Btn_GoToOutputRow()
    Dim nm As String, inputRow As Long
    On Error Resume Next
    nm = CStr(Application.Caller)
    On Error GoTo 0

    ' 도형 이름은 "goto_13" 처럼 입력정보 행 번호를 달고 있다
    If Left$(nm, 5) = "goto_" Then
        inputRow = CLng(Val(Mid$(nm, 6)))
    Else
        inputRow = ActiveCell.Row
    End If

    If Not IsContractRow(inputRow) Then
        MsgBox "이 줄에는 계약이 없습니다. 계약명을 먼저 넣어 주세요.", _
               vbInformation, "출력관리로 이동"
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = WsOutput()
    Dim r As Long: r = OutRow(inputRow)

    ws.Activate
    Application.Goto ws.Cells(r, OC_계약명), False
    ws.Cells(r, OC_계약명).Select

    ' 이 계약만 켠다
    Application.EnableEvents = False
    Dim i As Long
    For i = OUT_FIRST_ROW To OUT_LAST_ROW
        ws.Cells(i, OC_선택).Value = (i = r)
    Next i
    Application.EnableEvents = True
    SyncCheckboxes

    HighlightOutputRow r
End Sub

' 대상 행을 연두색으로 칠하고 5초 뒤 되돌린다.
'
' Interior.Color 로 칠하면 **보이지 않는다.** 이 시트에는 계약의 진행/완료를
' 칠하는 조건부서식이 늘 걸려 있고, 조건부서식은 언제나 셀 서식을 이기기 때문이다.
' 그래서 강조도 조건부서식으로 걸고, 맨 위 우선순위로 올린다.
Public Sub HighlightOutputRow(ByVal r As Long)
    ClearOutputHighlight                       ' 앞의 강조가 남아 있으면 먼저 지운다

    Dim ws As Worksheet: Set ws = WsOutput()
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, OC_요약))

    On Error Resume Next
    With rng.FormatConditions.Add(xlExpression, , "=TRUE")
        .SetFirstPriority                      ' 진행/완료 규칙보다 위로
        .Interior.Color = HIGHLIGHT_COLOR
        .StopIfTrue = True
    End With
    On Error GoTo 0

    mHighlightRow = r

    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 5), "ClearOutputHighlight"
    On Error GoTo 0
End Sub

' 강조로 걸어 둔 조건부서식만 떼어 낸다.
'
' 그 행의 조건부서식을 통째로 지우면 진행/완료 규칙까지 날아가고, 그 줄만
' 영영 색이 안 나온다. 수식이 "=TRUE" 인 것 ― 강조가 만든 그것 ― 만 찾아 지운다.
Public Sub ClearOutputHighlight()
    If mHighlightRow = 0 Then Exit Sub

    On Error Resume Next
    Dim ws As Worksheet: Set ws = WsOutput()
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(mHighlightRow, 1), ws.Cells(mHighlightRow, OC_요약))

    Dim i As Long
    For i = rng.FormatConditions.Count To 1 Step -1
        If rng.FormatConditions(i).Type = xlExpression Then
            If UCase$(rng.FormatConditions(i).Formula1) = "=TRUE" Then
                rng.FormatConditions(i).Delete
            End If
        End If
    Next i
    On Error GoTo 0

    mHighlightRow = 0
End Sub
