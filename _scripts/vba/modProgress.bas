Attribute VB_Name = "modProgress"
Option Explicit

'==============================================================================
' 계약 진행 / 완료 자동 판정
'
' 화면에서 계약 줄이 무슨 색인지는 **비고(T열) 값 하나**가 정한다. 색칠은
' 조건부서식이 하고(build_workbook.py), 이 모듈은 그 비고 칸을 채운다.
' 색의 출처를 하나로 묶어 두어야 '사용자가 드롭다운으로 바꾸면 그 값이 이긴다' 가
' 따로 손대지 않아도 성립한다.
'
'   계약 종료일이 지났으면   → "완료"  (진한 회색)
'   아직 남아 있으면         → "진행"  (연한 초록)
'   사용자가 직접 고르면     → 날짜와 무관하게 그 값 그대로
'
' 사용자가 직접 골랐다는 사실은 숨겨 둔 V열(CC_비고수동)에 적어 둔다. 색으로
' 표시할 수가 없기 때문이다 ― 줄 전체가 이미 진행/완료 색으로 덮여 있다.
'==============================================================================

Public Const REMARK_PROGRESS As String = "진행"
Public Const REMARK_DONE     As String = "완료"
Public Const REMARK_MANUAL   As String = "수동"      ' V열에 적히는 표시

' 한 줄의 비고를 날짜에 맞춰 다시 채운다. 사용자가 손으로 정한 줄은 건드리지 않는다.
Public Sub RefreshProgress(ByVal inputRow As Long)
    If inputRow < IN_FIRST_ROW Or inputRow > IN_LAST_ROW Then Exit Sub

    Dim ws As Worksheet: Set ws = WsInput()
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error GoTo Done

    ' 계약이 지워진 줄이면 비고도 표시도 함께 지운다. 남겨 두면 빈 줄에
    ' 색만 칠해져 계약이 있는 것처럼 보인다.
    If Trim$(CStr(ws.Cells(inputRow, CC_계약명).Value)) = "" Then
        ws.Cells(inputRow, CC_비고).ClearContents
        ws.Cells(inputRow, CC_비고수동).ClearContents
        GoTo Done
    End If

    If IsManualRemark(inputRow) Then GoTo Done

    Dim e As Variant
    e = ws.Cells(inputRow, CC_종료일).Value
    If Not IsDate(e) Then
        ws.Cells(inputRow, CC_비고).ClearContents
        GoTo Done
    End If

    If CDate(e) < Date Then
        ws.Cells(inputRow, CC_비고).Value = REMARK_DONE
    Else
        ws.Cells(inputRow, CC_비고).Value = REMARK_PROGRESS
    End If

Done:
    Application.EnableEvents = prevEvents
End Sub

' 100건 전부. 통합문서를 열 때 부른다 ― 파일을 덮어 둔 사이에 종료일이 지났으면
' 그때 저절로 완료로 넘어간다.
Public Sub RefreshAllProgress()
    Dim r As Long
    For r = IN_FIRST_ROW To IN_LAST_ROW
        RefreshProgress r
    Next r
End Sub

' 사용자가 비고를 직접 골랐는가
Public Function IsManualRemark(ByVal inputRow As Long) As Boolean
    IsManualRemark = (Trim$(CStr(WsInput().Cells(inputRow, CC_비고수동).Value)) = REMARK_MANUAL)
End Function

' 비고 칸이 사람 손으로 바뀌었을 때 (Sheet_입력정보 이벤트가 부른다)
'
' 값을 골랐으면 '수동' 으로 못박고, 칸을 비우면 표시를 지워 날짜 판정으로 되돌린다.
' 비우는 것이 곧 '자동으로 되돌리기' 다 ― 따로 버튼을 두지 않아도 된다.
Public Sub MarkRemarkManual(ByVal cel As Range)
    Dim ws As Worksheet: Set ws = WsInput()
    Dim r As Long: r = cel.Row
    If r < IN_FIRST_ROW Or r > IN_LAST_ROW Then Exit Sub

    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next

    If Trim$(CStr(cel.Value)) = "" Then
        ws.Cells(r, CC_비고수동).ClearContents
        Application.EnableEvents = prevEvents
        RefreshProgress r
        Exit Sub
    End If

    ws.Cells(r, CC_비고수동).Value = REMARK_MANUAL

    On Error GoTo 0
    Application.EnableEvents = prevEvents
End Sub

' 손으로 정한 것을 모두 풀고 날짜 기준으로 되돌린다. (필요하면 버튼에 걸 수 있다)
Public Sub ResetProgressToAuto()
    Dim ws As Worksheet: Set ws = WsInput()
    Dim prevEvents As Boolean
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    ws.Range(ws.Cells(IN_FIRST_ROW, CC_비고수동), _
             ws.Cells(IN_LAST_ROW, CC_비고수동)).ClearContents
    Application.EnableEvents = prevEvents

    RefreshAllProgress
End Sub
