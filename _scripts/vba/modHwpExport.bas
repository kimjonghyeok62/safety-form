Attribute VB_Name = "modHwpExport"
Option Explicit

'==============================================================================
' 한글(HWP) 출력 엔진
'
' 흐름:  템플릿(form\templates\*.hwpx)  →  {{태그}} 치환  →  임시폴더에 저장
'        →  필요한 것들을 한 문서로 이어붙임  →  한글 창에 띄움
'
' 템플릿 원본은 열기만 하고 절대 저장하지 않는다. 치환 결과는 항상
' %TEMP%\safetyform\ 아래 새 파일로 나가므로 실수로 서식을 덮어쓸 수 없다.
'==============================================================================

' 확인 / 확인·취소 / 예·아니오·취소 / 예·아니오 / 재시도 팝업에 긍정 응답 자동 클릭.
' 걸어두지 않으면 한글이 보이지 않는 곳에서 대화상자를 띄우고 멈춘다.
' &H111011 = 확인(&H1) + 확인/취소의 확인(&H10) + 예/아니오/취소의 예(&H1000)
'          + 예/아니오의 예(&H10000) + 재시도/취소의 재시도(&H100000)
Private Const MSGBOX_AUTO As Long = &H111011

' 여러 줄 값을 넣을 때 쓰는 구분자. 한글 찾아바꾸기는 줄바꿈 문자를 그대로
' 넣지 못하므로 가운뎃점으로 이어 붙인다.
Private Const LINE_SEP As String = "  ·  "

' 조용히 실패하면 원인을 알 수 없다. 마지막 오류를 담아 두었다가 끝에 한 번에 알린다.
Public LastError As String

' 무인 실행(점검 스크립트 등)에서 알림창이 뜨면 거기서 멈춰 버린다. 그럴 때 끄고 쓴다.
Public SilentMode As Boolean

' 마지막 병합 결과 ― 서식마다 새 쪽에서 시작했는지 점검할 때 쓴다.
' (VBA 는 모듈 수준 변수를 반드시 첫 프로시저보다 위에서 선언해야 한다)
Public LastMergedPages As Long
Public LastMergedFiles As Long
Public LastBreakFails As Long          ' 쪽 나눔이 먹지 않은 횟수 (0 이어야 한다)

' 버튼 한 번 동안만 들고 있는 것들. 쓰는 곳은 아래 TemplateTagSet 과
' BuildContentBlock 이고, ResetExportCache 가 비운다.
Private TagSetCache As Object       ' 서식 이름 → 그 서식에 든 태그 (시트와 무관해 안 비운다)
Private LibCache As Variant         ' ContentLibrary 시트를 뜬 배열
Private BlockCache As Object        ' (작업유형, 구분) → 이어 붙인 문구

Public Sub SetSilentMode(ByVal onOff As Boolean)
    SilentMode = onOff
End Sub

Public Function GetLastExportError() As String
    GetLastExportError = LastError
End Function

Public Function GetLastMergeInfo() As String
    GetLastMergeInfo = LastMergedFiles & "|" & LastMergedPages & "|" & LastBreakFails
End Function

' 점검용 ― BuildTagMap 이 채울 수 있는 태그 이름을 모두 돌려준다.
' verify_templates.py 가 이것과 서식 파일의 {{태그}} 를 맞대어 본다.
Public Function GetTagList(ByVal inputRow As Long) As String
    Dim d As Object, k As Variant, acc As String
    Set d = BuildTagMap(inputRow)
    For Each k In d.Keys
        If acc <> "" Then acc = acc & "|"
        acc = acc & CStr(k)
    Next k
    GetTagList = acc
End Function

' 점검용 ― 태그 하나가 어떤 값으로 채워지는지 돌려준다.
Public Function GetTagValue(ByVal inputRow As Long, ByVal tag As String) As String
    Dim d As Object
    Set d = BuildTagMap(inputRow)
    If d.Exists(tag) Then GetTagValue = CStr(d(tag))
End Function

Private Sub Notify(ByVal msg As String, ByVal style As Long, ByVal title As String)
    If SilentMode Then
        LastError = title & ": " & msg
    Else
        MsgBox msg, style, title
    End If
End Sub

Private Function ErrorSuffix() As String
    If LastError <> "" Then ErrorSuffix = vbLf & vbLf & "(" & LastError & ")"
End Function

'==============================================================================
' 버튼에서 부르는 진입점
'
' **버튼 하나하나가 오류의 마지막 방벽이다.** 여기까지 처리되지 않은 오류가
' 올라오면 엑셀은 VBA 편집기를 중단 모드로 띄운 채 멈춘다. 통합문서를 보이지 않게
' 열어 둔 점검 중이면 그 창이 떴다 사라지는 것처럼 보이고, 사용자가 쓰는 중이면
' 코드 창이 튀어나온다. 어느 쪽도 있어서는 안 된다.
'
' 그래서 진입점마다 GuardFailed 로 받아 안내창 한 장으로 끝낸다. 원인은
' LastError 에 남으므로 GetLastExportError 로 읽을 수 있다.
'==============================================================================

' 진입점에서 오류를 받아 마무리한다. 안내창을 걷고, 무슨 일이었는지 남긴다.
Private Sub GuardFailed(ByVal caption As String)
    Dim errNo As Long, errMsg As String
    errNo = Err.Number
    errMsg = Err.Description
    If Trim$(errMsg) = "" Then errMsg = "(설명 없음)"

    On Error Resume Next
    HideNotice
    On Error GoTo 0

    If LastError = "" Then
        LastError = caption & " 처리 중 오류 " & errNo & ": " & errMsg
    End If

    Notify "출력을 마치지 못했습니다." & vbLf & vbLf & _
           "한글이 도중에 닫혔거나 응답하지 않을 때 이렇게 됩니다." & vbLf & _
           "한글을 모두 닫고 다시 눌러 주세요." & ErrorSuffix(), _
           vbExclamation, caption
End Sub

' [1~7번 출력]
Public Sub Btn_ExportHwpForms()
    On Error GoTo Failed
    Dim rows As Variant
    rows = CheckedInputRows()
    If UBound(rows) < 0 Then
        Notify "출력할 계약이 없습니다." & vbLf & _
               "왼쪽 체크박스로 계약을 고르거나, 계약이 있는 줄에 커서를 두고 누르세요.", _
               vbInformation, "선택된 계약 없음"
        Exit Sub
    End If

    ' 시트에서 떠 온 것을 버리고 다시 읽는다. 안내문을 만들 때 이미 사전을 쓰므로
    ' 반드시 ShowNotice 보다 앞이어야 한다.
    ResetExportCache

    ' 무엇을 몇 종 만드는지 먼저 띄운다. 한글이 뜰 때까지 몇십 초가 걸리므로
    ' 이 창이 없으면 버튼이 눌린 것인지도 알 수 없다.
    ShowNotice "1~7번 출력", NoticeText17(rows)

    Dim hwp As Object
    Set hwp = HwpConnect()
    If hwp Is Nothing Then
        HideNotice
        Exit Sub
    End If
    LastError = ""

    ' 계약을 모두 돌며 **경로만** 모으고, 병합은 맨 끝에 딱 한 번 한다.
    '
    ' 예전에는 계약마다 병합까지 했는데, 병합은 hwp.Open 으로 시작한다.
    ' 같은 한글 인스턴스에서 Open 을 다시 부르면 앞 문서가 **교체**되므로
    ' 두 건을 고르면 마지막 계약만 남았다. 경로를 다 모아 한 번에 이어붙이면
    ' 계약과 계약 사이에도 쪽 나눔이 들어가 계약별로 새 쪽에서 시작한다.
    Dim paths() As String, n As Long
    ReDim paths(0 To (UBound(rows) - LBound(rows) + 1) * FORM_MERGE_LAST)
    n = 0

    Dim madeAny As Boolean, skipped As String
    Dim i As Long, before As Long
    For i = LBound(rows) To UBound(rows)
        before = n
        RenderContractForms hwp, CLng(rows(i)), paths, n
        If n > before Then
            madeAny = True
        Else
            skipped = skipped & vbLf & " · " & WsInput().Cells(rows(i), CC_계약명).Value
        End If
    Next i

    If n > 0 Then
        ReDim Preserve paths(0 To n - 1)
        ' 1~7 묶음은 계약 하나당 여러 서식이 한 문서로 붙으므로 '중대재해예방' 으로 낸다.
        ' Function 은 반드시 결과를 받아 부른다 ― 문장처럼 부르면 한셀이 컴파일을 멈춘다.
        Dim ok As Boolean
        ok = MergeIntoNewDoc(hwp, paths, _
                             OutName(CLng(rows(LBound(rows))), "중대재해예방", _
                                     UBound(rows) - LBound(rows)))
    End If

    ShowHwp hwp
    HideNotice

    If Not madeAny Then
        Notify "선택한 계약에서 만들어진 서식이 없습니다." & vbLf & _
               "출력관리 시트에 ○ 표시가 있는지 확인하세요." & ErrorSuffix(), _
               vbExclamation, "출력할 서식 없음"
    ElseIf skipped <> "" Then
        Notify "아래 계약은 필요한 서식이 없어 건너뛰었습니다:" & skipped & ErrorSuffix(), _
               vbInformation, "일부 건너뜀"
    End If
    Exit Sub
Failed:
    GuardFailed "1~7번 출력"
End Sub

' [8. 안전보건협의체 회의록] ― 달마다 1장(그 달의 마지막 순회점검일)
Public Sub Btn_ExportCouncil()
    On Error GoTo Failed
    ExportCouncil
    Exit Sub
Failed:
    GuardFailed "안전보건협의체 회의록"
End Sub

' [순회점검표 출력] ― 계약기간을 달 단위로 쪼개 달마다 1장
Public Sub Btn_ExportPatrol()
    On Error GoTo Failed
    ExportPatrol
    Exit Sub
Failed:
    GuardFailed "순회점검표"
End Sub

' [합동점검표 출력] ― 90일 구간마다 마지막 지정 요일에 1장(분기 1회)
Public Sub Btn_ExportJoint()
    On Error GoTo Failed
    ExportRepeating "10_합동점검표.hwpx", True, CriteriaJoint(), "합동점검표"
    Exit Sub
Failed:
    GuardFailed "합동점검표"
End Sub

' [11~15] 회색 버튼 다섯 개 ― 계약 1건에 1부씩
Public Sub Btn_Export11()
    On Error GoTo Failed
    ExportSingleForm 11, "아차사고 사례 제안서"
    Exit Sub
Failed:
    GuardFailed "아차사고 사례 제안서"
End Sub

Public Sub Btn_Export12()
    On Error GoTo Failed
    ExportSingleForm 12, "관리감독자 업무수행 평가표"
    Exit Sub
Failed:
    GuardFailed "관리감독자 업무수행 평가표"
End Sub

Public Sub Btn_Export13()
    On Error GoTo Failed
    ExportSingleForm 13, "중대재해 발생보고서"
    Exit Sub
Failed:
    GuardFailed "중대재해 발생보고서"
End Sub

Public Sub Btn_Export14()
    On Error GoTo Failed
    ExportSingleForm 14, "중대재해 사고보고서"
    Exit Sub
Failed:
    GuardFailed "중대재해 사고보고서"
End Sub

Public Sub Btn_Export15()
    On Error GoTo Failed
    ExportSingleForm 15, "계약이행 통합 서약서"
    Exit Sub
Failed:
    GuardFailed "계약이행 통합 서약서"
End Sub

'==============================================================================
' 계약 1건 → 서식 1~7 임시본 만들기
'
' 여기서는 **병합하지 않는다.** 경로만 paths 에 쌓아 돌려주고, 병합은 부르는 쪽이
' 고른 계약을 다 돈 뒤 한 번만 한다. (Btn_ExportHwpForms 의 주석 참고)
'==============================================================================

Private Sub RenderContractForms(ByVal hwp As Object, ByVal inputRow As Long, _
                                ByRef paths() As String, ByRef n As Long)
    Dim wsOut As Worksheet: Set wsOut = WsOutput()
    Dim oRow As Long: oRow = OutRow(inputRow)

    Dim tags As Object
    Set tags = BuildTagMap(inputRow)

    Dim formNo As Long, tmpl As String, rendered As String
    For formNo = 1 To FORM_MERGE_LAST
        If IsMarked(wsOut.Cells(oRow, FormCol(formNo)).Value) Then
            tmpl = TemplateFileFor(formNo, inputRow)
            If tmpl <> "" Then
                rendered = RenderTemplate(hwp, tmpl, tags, _
                                          "계약" & inputRow & "_서식" & formNo)
                If rendered <> "" Then
                    paths(n) = rendered
                    n = n + 1
                End If
            End If
        End If
    Next formNo
End Sub

' 서식 번호 → 템플릿 파일명 (1번은 계약 성격에 따라 셋 중 하나)
'
' 한글 문서 안에 찍힌 '서식 6' 같은 번호는 경기도교육청 매뉴얼의 고유번호라
' 여기 번호(1~15)와 다르다. 그래도 맞다 ― 파일명과 엑셀 열 번호만 맞으면 된다.
'
' **파일 이름 앞자리 = 화면 번호 = 출력관리 열 차례**로 맞춰 두었다. 협의체
' 회의록을 3번에서 8번으로 옮기면서 파일도 함께 옮겨 이름을 다시 매겼다
' (03_협의체회의록 → 08_협의체회의록, 04~08 → 03~07). 셋이 어긋나 있으면
' 폴더를 열었을 때 어느 것이 화면의 몇 번인지 매번 되짚어야 한다.
Private Function TemplateFileFor(ByVal formNo As Long, ByVal inputRow As Long) As String
    Select Case formNo
        Case 1:  TemplateFileFor = ResolvePlanTemplate(inputRow)
        Case 2:  TemplateFileFor = "02_수준평가표.hwpx"
        Case 3:  TemplateFileFor = "03_도급점검표.hwpx"
        Case 4:  TemplateFileFor = "04_점검표_고소.hwpx"
        Case 5:  TemplateFileFor = "05_점검표_화기.hwpx"
        Case 6:  TemplateFileFor = "06_점검표_밀폐.hwpx"
        Case 7:  TemplateFileFor = "07_점검표_전기.hwpx"
        Case 8:  TemplateFileFor = "08_협의체회의록.hwpx"
        Case 9:  TemplateFileFor = "09_순회점검표.hwpx"
        Case 10: TemplateFileFor = "10_합동점검표.hwpx"
        Case 11: TemplateFileFor = "11_아차사고제안서.hwpx"
        Case 12: TemplateFileFor = "12_관리감독자평가표.hwpx"
        Case 13: TemplateFileFor = "13_중대재해발생보고서.hwpx"
        Case 14: TemplateFileFor = "14_중대재해사고보고서.hwpx"
        Case 15: TemplateFileFor = "15_계약이행서약서.hwpx"
        Case Else: TemplateFileFor = ""
    End Select
End Function

'==============================================================================
' 11~15 ― 서식 하나를 계약마다 1부씩
'==============================================================================

Private Sub ExportSingleForm(ByVal formNo As Long, ByVal caption As String)
    Dim rows As Variant
    rows = CheckedInputRows()
    If UBound(rows) < 0 Then
        Notify "출력할 계약이 없습니다." & vbLf & _
               "왼쪽 체크박스로 계약을 고르거나, 계약이 있는 줄에 커서를 두고 누르세요.", _
               vbInformation, caption
        Exit Sub
    End If

    ResetExportCache

    Dim hwp As Object
    Set hwp = HwpConnect()
    If hwp Is Nothing Then Exit Sub
    LastError = ""

    Dim tmpl As String
    tmpl = TemplateFileFor(formNo, CLng(rows(LBound(rows))))

    Dim paths() As String, n As Long
    ReDim paths(0 To UBound(rows) - LBound(rows))
    n = 0

    Dim i As Long, rendered As String
    For i = LBound(rows) To UBound(rows)
        rendered = RenderTemplate(hwp, tmpl, BuildTagMap(CLng(rows(i))), _
                                  "계약" & rows(i) & "_서식" & formNo)
        If rendered <> "" Then
            paths(n) = rendered
            n = n + 1
        End If
    Next i

    If n = 0 Then
        Notify caption & "를 만들지 못했습니다." & ErrorSuffix(), vbExclamation, caption
        Exit Sub
    End If

    ReDim Preserve paths(0 To n - 1)
    Dim ok As Boolean
    ok = MergeIntoNewDoc(hwp, paths, _
                         OutName(CLng(rows(LBound(rows))), LabelOf(tmpl), _
                                 UBound(rows) - LBound(rows)))
    ShowHwp hwp
    If Not ok Then
        Notify caption & "를 이어붙이지 못했습니다." & ErrorSuffix(), vbExclamation, caption
    End If
End Sub

'==============================================================================
' 순회점검표 ― 달마다 1장
'
' 신 양식은 한 장에 1주~5주 행이 들어 있는 '한 달치' 서식이다.
' 그래서 매주 1장이 아니라 **달마다 1장**을 만들고, 그 달의 점검일을
' 1주~5주 칸에 나눠 넣는다. 1년 계약이면 52장이 아니라 12장이다.
'==============================================================================

Private Sub ExportPatrol()
    Dim rows As Variant
    rows = CheckedInputRows()
    If UBound(rows) < 0 Then
        Notify "출력할 계약이 없습니다." & vbLf & _
               "왼쪽 체크박스로 계약을 고르거나, 계약이 있는 줄에 커서를 두고 누르세요.", _
               vbInformation, "순회점검표"
        Exit Sub
    End If

    ResetExportCache
    ShowNotice "9. 순회점검표 출력", NoticeTextRepeat(rows, False, CriteriaPatrol())

    Dim hwp As Object
    Set hwp = HwpConnect()
    If hwp Is Nothing Then
        HideNotice
        Exit Sub
    End If
    LastError = ""

    Dim notes As String, madeAny As Boolean
    Dim i As Long, inputRow As Long, nm As String
    For i = LBound(rows) To UBound(rows)
        inputRow = CLng(rows(i))
        nm = CStr(WsInput().Cells(inputRow, CC_계약명).Value)

        If ContractDays(inputRow) < CriteriaPatrol() Then
            notes = notes & vbLf & " · " & nm & " ― 계약기간 " & ContractDays(inputRow) & _
                    "일 (" & CriteriaPatrol() & "일 이상일 때 필요)"
        Else
            Dim months As Collection
            Set months = PatrolMonths(inputRow)
            If months.Count = 0 Then
                notes = notes & vbLf & " · " & nm & " ― 계약기간 안에 " & _
                        CStr(WsInput().Cells(inputRow, CC_점검요일).Value) & "이 없습니다"
            Else
                BuildPatrolDoc hwp, inputRow, months
                madeAny = True
            End If
        End If
    Next i

    ShowHwp hwp
    HideNotice

    If notes <> "" Then
        Notify "순회점검표를 만들지 않은 계약이 있습니다:" & notes & ErrorSuffix(), _
               vbInformation, "순회점검표"
    ElseIf Not madeAny Then
        Notify "만들어진 순회점검표가 없습니다." & ErrorSuffix(), vbInformation, "순회점검표"
    End If
End Sub

' 계약기간 안의 점검일을 달 단위로 묶는다.
'
' 각 항목 = Array(년, 월, 점검일1…점검일5, 요일1…요일5) ― 없는 주차는 빈 문자열.
' (한 달에 같은 요일은 많아야 다섯 번 나온다)
'
' 날짜와 요일을 **따로** 담는다. 순회점검표의 점검일 칸이 17mm 밖에 안 되어
' '2026.6.4.(목)' 이 한 줄에 안 들어가는데, 한 덩어리로 넣으면 여는 괄호까지만
' 날짜 줄에 남고 '목)' 만 다음 줄로 넘어간다. 서식 쪽에 줄을 하나 더 만들어 두고
' 요일은 그 줄에 넣는다.
Public Function PatrolMonths(ByVal inputRow As Long) As Collection
    Dim out As New Collection
    Dim dates As Variant
    dates = InspectionDates(inputRow, False)
    If UBound(dates) < 0 Then
        Set PatrolMonths = out
        Exit Function
    End If

    Dim cur As Variant, curYm As String, ym As String
    Dim slot As Long, i As Long, d As Date

    For i = LBound(dates) To UBound(dates)
        d = dates(i)
        ym = Format$(d, "yyyymm")
        If ym <> curYm Then
            If curYm <> "" Then out.Add cur
            cur = Array(Format$(d, "yyyy"), CStr(Month(d)), _
                        "", "", "", "", "", "", "", "", "", "")
            curYm = ym
            slot = 0
        End If
        If slot < 5 Then
            cur(2 + slot) = FmtInspectDate(d)
            cur(7 + slot) = FmtDow(d)
            slot = slot + 1
        End If
    Next i
    If curYm <> "" Then out.Add cur

    Set PatrolMonths = out
End Function

Private Sub BuildPatrolDoc(ByVal hwp As Object, ByVal inputRow As Long, _
                           ByVal months As Collection)
    Dim tags As Object
    Set tags = BuildTagMap(inputRow)

    Dim paths() As String
    ReDim paths(0 To months.Count - 1)

    Dim n As Long, i As Long, m As Variant, k As Long
    n = 0
    For i = 1 To months.Count
        m = months(i)
        tags("{{점검년}}") = CStr(m(0))
        tags("{{점검월}}") = CStr(m(1))
        For k = 1 To 5
            tags("{{점검일" & k & "}}") = CStr(m(1 + k))
            tags("{{점검요일" & k & "}}") = CStr(m(6 + k))
        Next k
        ' 그 달의 첫 점검일을 대표 점검일로 둔다(다른 태그를 쓰는 서식과 맞추기 위해).
        ' 이쪽은 한 줄짜리 자리라 요일을 붙여 쓴다.
        tags("{{점검일자}}") = CStr(m(2)) & CStr(m(7))
        tags("{{회차}}") = i & "회차"

        paths(n) = RenderTemplate(hwp, "09_순회점검표.hwpx", tags, _
                                  "계약" & inputRow & "_순회_" & Format$(i, "00"))
        If paths(n) <> "" Then n = n + 1
    Next i

    If n = 0 Then Exit Sub
    ReDim Preserve paths(0 To n - 1)
    Dim ok As Boolean
    ok = MergeIntoNewDoc(hwp, paths, OutName(inputRow, "순회점검표_총" & n & "회차"))
End Sub

'==============================================================================
' 안전보건협의체 회의록 ― 달마다 1장 (월 1회)
'
' 협의체는 월 1회 이상 열어야 한다. 회의는 어차피 현장을 도는 날에 맞춰 열리므로
' **그 달의 마지막 순회점검일**을 회의일로 잡는다. 달의 끝에 여는 편이 그달에
' 있었던 일을 다 담을 수 있고, 순회점검·합동점검과 같은 요일이라 헛걸음이 없다.
'
' 서식이 매번 같아 회차 구분이 안 되면 열두 장이 한 장처럼 보인다. 그래서
' ApplyCouncilRound 가 회차마다 안건·위험요인·의결사항을 조금씩 바꿔 넣는다.
'==============================================================================

' 계약기간 안의 순회점검일을 달로 묶어 **달마다 마지막 하나**만 남긴다.
' InspectionDates 의 quarterlyOnly 처리와 같은 틀이다 ― 같은 달이면 앞서 넣은
' 것을 덮어쓰고, 달이 바뀔 때 비로소 새 자리에 넣는다.
Public Function CouncilDates(ByVal inputRow As Long) As Variant
    Dim dates As Variant
    dates = InspectionDates(inputRow, False)
    If UBound(dates) < 0 Then
        CouncilDates = Array()
        Exit Function
    End If

    Dim acc() As Date, n As Long
    ReDim acc(0 To UBound(dates) - LBound(dates))
    n = 0

    Dim curYm As String, ym As String, i As Long, d As Date
    For i = LBound(dates) To UBound(dates)
        d = CDate(dates(i))
        ym = Format$(d, "yyyymm")
        If ym = curYm And n > 0 Then
            acc(n - 1) = d                  ' 같은 달 ― 마지막 날로 밀어 둔다
        Else
            acc(n) = d
            n = n + 1
            curYm = ym
        End If
    Next i

    ReDim Preserve acc(0 To n - 1)
    CouncilDates = acc
End Function

Private Sub ExportCouncil()
    Dim rows As Variant
    rows = CheckedInputRows()
    If UBound(rows) < 0 Then
        Notify "출력할 계약이 없습니다." & vbLf & _
               "왼쪽 체크박스로 계약을 고르거나, 계약이 있는 줄에 커서를 두고 누르세요.", _
               vbInformation, "협의체 회의록"
        Exit Sub
    End If

    ResetExportCache
    ShowNotice "8. 안전보건협의체 회의록 출력", NoticeTextCouncil(rows)

    Dim hwp As Object
    Set hwp = HwpConnect()
    If hwp Is Nothing Then
        HideNotice
        Exit Sub
    End If
    LastError = ""

    Dim notes As String, madeAny As Boolean
    Dim i As Long, inputRow As Long, nm As String, minDays As Long
    minDays = CriteriaCouncil()

    For i = LBound(rows) To UBound(rows)
        inputRow = CLng(rows(i))
        nm = CStr(WsInput().Cells(inputRow, CC_계약명).Value)

        If ContractDays(inputRow) < minDays Then
            notes = notes & vbLf & " · " & nm & " ― 계약기간 " & ContractDays(inputRow) & _
                    "일 (" & minDays & "일 이상일 때 필요)"
        Else
            Dim dates As Variant
            dates = CouncilDates(inputRow)
            If UBound(dates) < 0 Then
                notes = notes & vbLf & " · " & nm & " ― 계약기간 안에 " & _
                        CStr(WsInput().Cells(inputRow, CC_점검요일).Value) & "이 없습니다"
            Else
                BuildCouncilDoc hwp, inputRow, dates
                madeAny = True
            End If
        End If
    Next i

    ShowHwp hwp
    HideNotice

    If notes <> "" Then
        Notify "협의체 회의록을 만들지 않은 계약이 있습니다:" & notes & ErrorSuffix(), _
               vbInformation, "협의체 회의록"
    ElseIf Not madeAny Then
        Notify "만들어진 협의체 회의록이 없습니다." & ErrorSuffix(), vbInformation, "협의체 회의록"
    End If
End Sub

Private Sub BuildCouncilDoc(ByVal hwp As Object, ByVal inputRow As Long, _
                            ByVal dates As Variant)
    Dim paths() As String
    ReDim paths(LBound(dates) To UBound(dates))

    Dim i As Long, n As Long, tags As Object
    n = 0
    For i = LBound(dates) To UBound(dates)
        ' 회차마다 사전을 새로 뜬다. 슬롯 태그를 회차 몫으로 갈아 끼우므로
        ' 한 벌을 돌려 쓰면 앞 회차 값이 뒤 회차에 새어 들어간다.
        Set tags = BuildTagMap(inputRow)
        ApplyCouncilRound tags, i + 1, CDate(dates(i))

        paths(n) = RenderTemplate(hwp, "08_협의체회의록.hwpx", tags, _
                                  "계약" & inputRow & "_협의체_" & Format$(i + 1, "00"))
        If paths(n) <> "" Then n = n + 1
    Next i

    If n = 0 Then Exit Sub
    ReDim Preserve paths(0 To n - 1)
    ' 사용자가 부르는 이름 그대로 낸다(파일 이름의 '협의체회의록' 보다 길지만
    ' 폴더에서 무슨 문서인지 한눈에 보이는 편이 낫다).
    Dim ok As Boolean
    ok = MergeIntoNewDoc(hwp, paths, _
                         OutName(inputRow, "안전보건협의체 회의록_총" & n & "회차"))
End Sub

' 회차 한 번 몫으로 태그를 갈아 끼운다.
'
' 무작위가 아니라 **돌려쓰기**다. 같은 계약을 다시 뽑아도 결과가 같아야
' 지난달에 낸 회의록과 이번에 낸 것이 어긋나지 않는다.
Private Sub ApplyCouncilRound(ByVal tags As Object, ByVal roundNo As Long, _
                              ByVal d As Date)
    tags("{{협의체회의일}}") = FmtDateDow(d)
    tags("{{회차}}") = roundNo & "회차"

    Dim off As Long: off = roundNo - 1

    ' 위험요인·안전대책 ― 사전에 든 항목을 회차만큼 밀어 가며 고른다
    tags("{{위험요인1}}") = RotateItem(CStr(tags("{{위험요인}}")), off + 1)
    tags("{{위험요인2}}") = RotateItem(CStr(tags("{{위험요인}}")), off + 2)
    tags("{{안전대책1}}") = RotateItem(CStr(tags("{{안전대책}}")), off + 1)

    ' 2번 연락방법 끝줄 ― 그 달의 계절·학사일정 안건. 표에 없는 달이면
    ' 어느 계약에서나 쓰는 상시 안건을 돌려 쓴다.
    Dim season As String
    season = SeasonAgenda(d)
    If season = "" Then
        tags("{{상시안건1}}") = RotateItem(CStr(tags("{{상시안건}}")), off + 1)
    Else
        tags("{{상시안건1}}") = season
    End If

    ' 6번 의결사항 ― 앞의 둘은 그 계약의 합의문에서, 마지막 하나는 회차 의결문에서.
    ' 합의문 세 줄이 매번 그대로면 회의록이 복사본으로 보인다.
    tags("{{협의체합의문1}}") = RotateItem(CStr(tags("{{협의체합의문}}")), off + 1)
    tags("{{협의체합의문2}}") = RotateItem(CStr(tags("{{협의체합의문}}")), off + 2)

    Dim resolution As String
    resolution = RoundResolution(roundNo)
    If resolution <> "" Then tags("{{협의체합의문3}}") = resolution
End Sub

' " · " 로 이은 항목 목록에서 n 번째. 끝을 넘으면 처음으로 돌아온다.
' (NthItem 은 없으면 "" 를 주지만 여기서는 늘 무언가가 나와야 한다)
Private Function RotateItem(ByVal s As String, ByVal n As Long) As String
    Dim cnt As Long
    cnt = ItemCount(s)
    If cnt = 0 Then Exit Function
    RotateItem = NthItem(s, ((n - 1) Mod cnt) + 1)
End Function

Private Function ItemCount(ByVal s As String) As Long
    If Trim$(s) = "" Then Exit Function

    Dim parts() As String, i As Long, k As Long
    parts = Split(s, " · ")
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then k = k + 1
    Next i
    ItemCount = k
End Function

' ContentLibrary 에서 그 달의 계절·학사일정 안건을 가져온다.
' 1열이 "월03" 처럼 달을 가리키고 2열이 "계절안건" 인 줄이다.
' 1열 값이 공통·고소작업 … 어느 것과도 겹치지 않아 BuildContentBlock 은 절대 걸리지 않는다.
Private Function SeasonAgenda(ByVal d As Date) As String
    On Error GoTo Fallback
    LoadLibCache

    Dim key As String, r As Long, acc As String
    key = "월" & Format$(d, "mm")
    For r = 2 To UBound(LibCache, 1)
        If Trim$(CStr(LibCache(r, 1) & "")) = key Then
            If Trim$(CStr(LibCache(r, 2) & "")) = "계절안건" Then
                If acc <> "" Then acc = acc & LINE_SEP
                acc = acc & CStr(LibCache(r, 4) & "")
            End If
        End If
    Next r
    SeasonAgenda = acc
    Exit Function
Fallback:
    SeasonAgenda = ""
End Function

' ContentLibrary 의 "회차의결" 줄을 회차 번호로 돌려 가며 하나 고른다.
Private Function RoundResolution(ByVal roundNo As Long) As String
    On Error GoTo Fallback
    LoadLibCache

    Dim acc() As String, n As Long, r As Long
    ReDim acc(0 To 200)
    For r = 2 To UBound(LibCache, 1)
        If Trim$(CStr(LibCache(r, 2) & "")) = "회차의결" Then
            If n <= UBound(acc) Then
                acc(n) = CStr(LibCache(r, 4) & "")
                n = n + 1
            End If
        End If
    Next r

    If n = 0 Then Exit Function
    RoundResolution = acc(((roundNo - 1) Mod n))
    Exit Function
Fallback:
    RoundResolution = ""
End Function

'==============================================================================
' 합동점검표 ― 회차만큼 반복 생성 (분기 1회)
'==============================================================================

Private Sub ExportRepeating(ByVal tmplName As String, ByVal quarterlyOnly As Boolean, _
                            ByVal minDays As Long, ByVal caption As String)
    Dim rows As Variant
    rows = CheckedInputRows()
    If UBound(rows) < 0 Then
        Notify "출력할 계약이 없습니다." & vbLf & _
               "왼쪽 체크박스로 계약을 고르거나, 계약이 있는 줄에 커서를 두고 누르세요.", _
               vbInformation, caption
        Exit Sub
    End If

    ResetExportCache
    ShowNotice "10. " & caption & " 출력", NoticeTextRepeat(rows, quarterlyOnly, minDays)

    Dim hwp As Object
    Set hwp = HwpConnect()
    If hwp Is Nothing Then
        HideNotice
        Exit Sub
    End If
    LastError = ""

    Dim notes As String, madeAny As Boolean
    Dim i As Long, inputRow As Long, nm As String
    For i = LBound(rows) To UBound(rows)
        inputRow = CLng(rows(i))
        nm = CStr(WsInput().Cells(inputRow, CC_계약명).Value)

        If ContractDays(inputRow) < minDays Then
            notes = notes & vbLf & " · " & nm & " ― 계약기간 " & ContractDays(inputRow) & _
                    "일 (" & minDays & "일 이상일 때 필요)"
        Else
            Dim dates As Variant
            dates = InspectionDates(inputRow, quarterlyOnly)
            If UBound(dates) < 0 Then
                notes = notes & vbLf & " · " & nm & " ― 계약기간 안에 해당 요일이 없습니다"
            Else
                BuildRepeatingDoc hwp, inputRow, tmplName, dates
                madeAny = True
            End If
        End If
    Next i

    ShowHwp hwp
    HideNotice

    If notes <> "" Then
        Notify caption & "를 만들지 않은 계약이 있습니다:" & notes & ErrorSuffix(), _
               vbInformation, caption
    ElseIf Not madeAny Then
        Notify "만들어진 " & caption & "가 없습니다." & ErrorSuffix(), vbInformation, caption
    End If
End Sub

Private Sub BuildRepeatingDoc(ByVal hwp As Object, ByVal inputRow As Long, _
                              ByVal tmplName As String, ByVal dates As Variant)
    Dim tags As Object
    Set tags = BuildTagMap(inputRow)

    Dim paths() As String
    ReDim paths(LBound(dates) To UBound(dates))

    Dim i As Long, n As Long
    n = 0
    For i = LBound(dates) To UBound(dates)
        tags("{{점검일자}}") = FmtDateDow(CDate(dates(i)))
        tags("{{회차}}") = (i + 1) & "회차"
        paths(n) = RenderTemplate(hwp, tmplName, tags, _
                                  "계약" & inputRow & "_" & Left$(tmplName, 2) & "_" & Format$(i + 1, "00"))
        If paths(n) <> "" Then n = n + 1
    Next i

    If n = 0 Then Exit Sub
    ReDim Preserve paths(0 To n - 1)
    Dim ok As Boolean
    ok = MergeIntoNewDoc(hwp, paths, _
                         OutName(inputRow, LabelOf(tmplName) & "_총" & n & "회차"))
End Sub

' 계약기간 안에서 지정 요일에 해당하는 날짜들.
' quarterlyOnly = True 이면 시작일로부터 90일씩 끊은 구간마다, 그 구간 안의
' **마지막** 순회점검일 하나만 남긴다(합동점검 = 분기 1회).
'
' 달로 끊지 않고 90일로 끊는 이유: 계약이 달 중간에 시작하면 첫 달과 끝 달이
' 반 토막이라, 달로 세면 3개월짜리 계약에서 합동점검이 4회가 되기도 한다.
' 시작일 기준 90일 구간은 계약이 언제 시작하든 분기 수가 어긋나지 않는다.
Public Function InspectionDates(ByVal inputRow As Long, ByVal quarterlyOnly As Boolean) As Variant
    Dim ws As Worksheet: Set ws = WsInput()
    Dim s As Variant, e As Variant
    s = ws.Cells(inputRow, CC_시작일).Value
    e = ws.Cells(inputRow, CC_종료일).Value
    If Not IsDate(s) Or Not IsDate(e) Then
        InspectionDates = Array()
        Exit Function
    End If

    Dim wd As Long
    wd = WeekdayNumber(CStr(ws.Cells(inputRow, CC_점검요일).Value))

    Dim acc() As Date, n As Long
    ReDim acc(0 To 500)
    n = 0

    ' 구간을 몇 개로 볼지 먼저 정한다.
    '
    ' 90 으로 나눈 몫을 그대로 구간 번호로 쓰면 끝에 며칠짜리 꼬리 구간이 생긴다.
    ' 365일 계약(90×4 + 5)에서 합동점검이 5회가 나오고 마지막 두 회차가 7일
    ' 간격으로 붙었다. 반올림해서 구간 수를 정하고, 꼬리는 마지막 구간이 삼킨다.
    Dim spanCount As Long, totalDays As Long
    totalDays = CLng(CDate(e)) - CLng(CDate(s)) + 1
    spanCount = Int(totalDays / 90 + 0.5)
    If spanCount < 1 Then spanCount = 1

    Dim lastPeriod As Long, period As Long
    lastPeriod = -1
    Dim d As Date, d2 As Date
    For d = CDate(s) To CDate(e)
        If Weekday(d) = wd Then
            period = (CLng(d) - CLng(CDate(s))) \ 90
            If period > spanCount - 1 Then period = spanCount - 1

            ' 그날이 공휴일이면 다음 평일로 미룬다. 계약기간을 넘어가면 그 회차는
            ' 버린다. 미룬 날이 앞 회차와 겹치는 일도 없어야 한다.
            d2 = NextWorkday(d)
            If d2 <= CDate(e) Then
                If quarterlyOnly And period = lastPeriod And n > 0 Then
                    ' 같은 90일 구간이면 앞서 넣은 것을 이 날로 덮어쓴다. 구간을
                    ' 다 훑고 나면 그 구간의 마지막 순회일만 남는다.
                    acc(n - 1) = d2
                ElseIf n <= UBound(acc) Then
                    If n = 0 Then
                        acc(n) = d2
                        n = n + 1
                    ElseIf acc(n - 1) <> d2 Then
                        acc(n) = d2
                        n = n + 1
                    End If
                End If
                lastPeriod = period
            End If
        End If
    Next d

    If n = 0 Then
        InspectionDates = Array()
    Else
        ReDim Preserve acc(0 To n - 1)
        InspectionDates = acc
    End If
End Function

'==============================================================================
' 한글 연결 / 치환 / 병합
'==============================================================================

Public Function HwpConnect() As Object
    Dim hwp As Object
    On Error Resume Next
    Set hwp = CreateObject("HWPFrame.HwpObject")
    On Error GoTo 0

    If hwp Is Nothing Then
        Notify "한글(HWP)을 실행할 수 없습니다." & vbLf & vbLf & _
               "· 이 컴퓨터에 한글이 설치되어 있는지 확인하세요." & vbLf & _
               "· 설치되어 있는데도 안 되면 설정 시트의 [보안모듈 등록] 버튼을 한 번 눌러 주세요.", _
               vbCritical, "한글 연결 실패"
        Exit Function
    End If

    On Error Resume Next
    hwp.RegisterModule "FilePathCheckDLL", "FilePathCheckerModule"
    hwp.SetMessageBoxMode MSGBOX_AUTO
    On Error GoTo 0

    Set HwpConnect = hwp
End Function

'==============================================================================
' 치환값 만들기
'
' 이름이 다른데 뜻이 같은 태그는 **모두 등록한다**. 서식마다 쓰는 이름이 조금씩
' 달라 하나만 등록하면 어떤 서식에서는 빈칸이 남는다.
'
' 결재란({{담당자}} {{관리자}} {{부서장}})에는 **직위만** 넣는다.
' 원본 서식의 결재란은 '○○ | ○○ | 학교장(부서장)' 처럼 직위를 적는 자리이고,
' 칸이 좁아 '직위 성명'을 넣으면 글자가 겹친다. 성명이 필요한 자리에는
' {{담당자성명}} 처럼 따로 쓴다.
'==============================================================================

Public Function BuildTagMap(ByVal inputRow As Long) As Object
    Dim ws As Worksheet: Set ws = WsInput()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    '--- 기관 기본정보 --------------------------------------------------------
    Dim sr As Long: sr = IN_SCHOOL_ROW
    Dim orgName As String
    orgName = CStr(ws.Cells(sr, IC_기관명).Value)

    d("{{기관구분}}") = CStr(ws.Cells(sr, IC_기관구분).Value)
    d("{{학교명}}") = orgName
    d("{{기관명}}") = orgName
    d("{{고유번호}}") = CStr(ws.Cells(sr, IC_고유번호).Value)
    d("{{주소}}") = CStr(ws.Cells(sr, IC_주소).Value)
    d("{{전화번호}}") = CStr(ws.Cells(sr, IC_전화번호).Value)
    d("{{팩스번호}}") = CStr(ws.Cells(sr, IC_팩스번호).Value)

    '--- 결재선 · 기관대표자 --------------------------------------------------
    Dim gr As Long: gr = IN_SIGN_ROW
    Dim dutyTitle As String, dutyName As String, dutyTel As String
    Dim mgrTitle As String, mgrName As String
    Dim headTitle As String, headName As String
    Dim ceoTitle As String, ceoName As String

    dutyTitle = CStr(ws.Cells(gr, SC_담당자직위).Value)
    dutyName = CStr(ws.Cells(gr, SC_담당자성명).Value)
    dutyTel = CStr(ws.Cells(gr, SC_담당자전화).Value)
    mgrTitle = CStr(ws.Cells(gr, SC_관리자직위).Value)
    mgrName = CStr(ws.Cells(gr, SC_관리자성명).Value)
    headTitle = CStr(ws.Cells(gr, SC_부서장직위).Value)
    headName = CStr(ws.Cells(gr, SC_부서장성명).Value)
    ceoTitle = CStr(ws.Cells(gr, SC_대표자직위).Value)
    ceoName = CStr(ws.Cells(gr, SC_대표자성명).Value)

    ' 결재란 = 직위
    d("{{담당자}}") = dutyTitle
    d("{{관리자}}") = mgrTitle
    d("{{부서장}}") = headTitle

    d("{{담당자직위}}") = dutyTitle
    d("{{담당자성명}}") = dutyName
    d("{{담당자전화번호}}") = dutyTel
    d("{{관리자직위}}") = mgrTitle
    d("{{관리자성명}}") = mgrName
    d("{{부서장직위}}") = headTitle
    d("{{부서장성명}}") = headName

    d("{{기관대표자}}") = JoinTitleName(ceoTitle, ceoName)
    d("{{기관대표자직위}}") = ceoTitle
    d("{{기관대표자성명}}") = ceoName

    ' 옛 이름 ― 사용자가 만든 서식이 이 이름을 쓰고 있을 수 있다
    d("{{행정실담당자}}") = dutyName
    d("{{행정실전화번호}}") = dutyTel
    d("{{계약담당자}}") = dutyName
    d("{{행정실장}}") = mgrName
    d("{{교장}}") = ceoName
    d("{{대표자}}") = ceoName        ' 기관 대표자. 업체 쪽은 {{업체대표자}} 를 쓴다

    '--- 비상연락망 · 관할기관 ------------------------------------------------
    Dim er As Long: er = IN_EMERG_ROW
    d("{{관할소방서}}") = CStr(ws.Cells(er, EC_소방서).Value)
    d("{{소방서}}") = CStr(ws.Cells(er, EC_소방서).Value)
    d("{{소방서전화번호}}") = CStr(ws.Cells(er, EC_소방서전화).Value)
    d("{{인근병원}}") = CStr(ws.Cells(er, EC_병원).Value)
    d("{{병원}}") = CStr(ws.Cells(er, EC_병원).Value)
    d("{{병원전화번호}}") = CStr(ws.Cells(er, EC_병원전화).Value)
    d("{{관할고용노동부}}") = CStr(ws.Cells(er, EC_노동지청).Value)
    d("{{관할고용노동지청}}") = CStr(ws.Cells(er, EC_노동지청).Value)
    d("{{고용노동부전화번호}}") = CStr(ws.Cells(er, EC_노동지청전화).Value)

    '--- 계약정보 -------------------------------------------------------------
    Dim vendorTel As String, staff As String, staffTel As String
    vendorTel = CStr(ws.Cells(inputRow, CC_대표자전화).Value)
    staff = CStr(ws.Cells(inputRow, CC_안전담당자).Value)
    staffTel = CStr(ws.Cells(inputRow, CC_담당자전화).Value)

    d("{{연번}}") = CStr(ws.Cells(inputRow, CC_연번).Value)
    d("{{계약명}}") = CStr(ws.Cells(inputRow, CC_계약명).Value)
    d("{{업체명}}") = CStr(ws.Cells(inputRow, CC_업체명).Value)
    d("{{사업자등록번호}}") = CStr(ws.Cells(inputRow, CC_사업자번호).Value)
    d("{{업체주소}}") = CStr(ws.Cells(inputRow, CC_업체주소).Value)
    d("{{업체대표자}}") = CStr(ws.Cells(inputRow, CC_업체대표).Value)
    d("{{대표자명}}") = CStr(ws.Cells(inputRow, CC_업체대표).Value)
    d("{{업체대표자연락처}}") = vendorTel
    d("{{업체전화}}") = vendorTel
    d("{{안전담당자}}") = staff
    d("{{업체담당자}}") = staff
    d("{{담당자전화}}") = staffTel
    d("{{안전담당자연락처}}") = staffTel
    d("{{업체담당자연락처}}") = staffTel
    d("{{점검요일}}") = CStr(ws.Cells(inputRow, CC_점검요일).Value)

    Dim s As Variant, e As Variant, days As Long
    s = ws.Cells(inputRow, CC_시작일).Value
    e = ws.Cells(inputRow, CC_종료일).Value
    days = ContractDays(inputRow)

    d("{{계약시작일}}") = FmtDate(s)
    d("{{계약종료일}}") = FmtDate(e)
    d("{{계약일수}}") = IIf(days > 0, days & "일간", "")
    If days > 0 Then
        d("{{계약기간}}") = FmtDate(s) & " ~ " & FmtDate(e) & "(" & days & "일간)"
    Else
        d("{{계약기간}}") = ""
    End If

    Dim amt As Variant
    amt = ws.Cells(inputRow, CC_계약금액).Value
    If IsNumeric(amt) And Trim$(CStr(amt)) <> "" Then
        d("{{계약금액}}") = Format$(amt, "#,##0") & "원"
        d("{{계약금액천원}}") = Format$(CDbl(amt) / 1000, "#,##0") & "천원"
    Else
        d("{{계약금액}}") = ""
        d("{{계약금액천원}}") = ""
    End If

    '--- 날짜 파생 -----------------------------------------------------------
    ' 날짜 하나를 짚는 태그에는 요일을 붙인다(2026.8.10.(월)). 계약기간 쪽은 안 붙인다.
    d("{{작성일}}") = FmtDateDow(Date)
    d("{{평가일}}") = FmtDateW(s)
    ' 협의체 회의일의 기본값은 계약 시작일(주말·공휴일이면 다음 평일)이다.
    d("{{협의체회의일}}") = FmtDateW(FirstWeekday(s))
    d("{{점검일자}}") = ""
    d("{{회차}}") = ""
    d("{{점검년}}") = ""
    d("{{점검월}}") = ""
    ' 순회점검표는 날짜와 요일을 두 줄로 나눠 넣는다(PatrolMonths 주석 참고).
    ' 값이 없어도 반드시 등록해야 문서에 태그 글자가 남지 않는다.
    Dim k As Long
    For k = 1 To 5
        d("{{점검일" & k & "}}") = ""
        d("{{점검요일" & k & "}}") = ""
    Next k

    '--- 계약유형 사전에서 가져오는 문구 --------------------------------------
    Dim types As String
    types = ActiveWorkTypes(inputRow)
    d("{{작업유형목록}}") = Replace(types, "|", ", ")

    Dim tname As String
    tname = Trim$(CStr(ws.Cells(inputRow, CC_계약유형).Value))
    d("{{계약유형}}") = tname
    d("{{작업내용}}") = ContractTypeField(tname, DICT_작업내용)
    d("{{협의체안건}}") = ContractTypeField(tname, DICT_협의안건)
    d("{{점검착안사항}}") = ContractTypeField(tname, DICT_점검착안)
    d("{{위험요인}}") = MergeBlocks(ContractTypeField(tname, DICT_위험요인), _
                                    BuildContentBlock(types, "위험요인"))
    d("{{안전대책}}") = MergeBlocks(ContractTypeField(tname, DICT_안전대책), _
                                    BuildContentBlock(types, "안전대책"))

    ' 계획서·회의록의 파란 예시 글을 대신하는 한 줄짜리 문구.
    ' 계약유형 사전에 걸리지 않은 계약은 빈칸이 아니라 일반적인 문구를 쓴다.
    ' 빈칸으로 두면 '- 작업절차:' 처럼 라벨만 남아 서식이 미완성으로 보인다.
    d("{{작업절차}}") = OrElse(ContractTypeField(tname, DICT_작업절차), _
        "작업 전 현장 확인 → 작업구역 통제 → 작업 시행 → 정리정돈 → 작업 후 확인")
    d("{{안전점검계획}}") = OrElse(ContractTypeField(tname, DICT_점검계획), _
        "작업장 주변 통제, 유해·위험 기계기구 점검")
    d("{{보호구방호장치}}") = OrElse(ContractTypeField(tname, DICT_보호구), _
        "안전모, 안전화, 보호장갑 등 개인보호구 지급 및 착용 확인")
    d("{{안전보건교육}}") = OrElse(ContractTypeField(tname, DICT_교육계획), _
        "작업 전 기계·기구 사용 시 주의사항 및 안전수칙 준수 교육")
    d("{{대피방법}}") = OrElse(ContractTypeField(tname, DICT_대피방법), _
        "작업을 즉시 멈추고 가까운 계단으로 건물 밖 안전한 곳까지 이동해 인원을 확인, " & _
        "다친 사람이 있으면 보건실을 거쳐 인근 병원으로 이송")

    ' 학교 현장 항목. 계약유형 사전에 걸리지 않은 계약도 빈칸으로 두지 않는다.
    d("{{작업권장시기}}") = OrElse(ContractTypeField(tname, DICT_권장시기), _
        "학사일정을 확인해 수업에 지장이 없는 시간대에 시행")
    d("{{위험성등급}}") = OrElse(ContractTypeField(tname, DICT_위험등급), "중")
    d("{{학교확인사항}}") = OrElse(ContractTypeField(tname, DICT_학교확인), _
        "작업 구역과 학생 통행로 분리 확인" & LINE_SEP & _
        "작업 일정과 범위를 관련 교직원에게 안내" & LINE_SEP & _
        "작업 후 정리 상태를 확인한 뒤 인수")
    d("{{업체조치사항}}") = OrElse(ContractTypeField(tname, DICT_업체조치), _
        "작업 전 안전교육을 하고 보호구 착용 확인" & LINE_SEP & _
        "작업 구역에 통제 표지를 세우고 학생 접근 차단" & LINE_SEP & _
        "작업 후 공구와 폐자재를 당일 치우고 보고")
    d("{{협의체합의문}}") = OrElse(ContractTypeField(tname, DICT_합의문), _
        "작업 구역과 학생 통행로를 분리하고 통제 표지를 두기로 합의함" & LINE_SEP & _
        "소음과 먼지가 나는 작업은 수업시간을 피해 시행하기로 협의함" & LINE_SEP & _
        "작업 후 정리 상태를 학교 담당자와 함께 확인하기로 합의함")
    d("{{상시안건}}") = BuildContentBlock(types, "상시안건")

    '--- 슬롯 태그 ― 항목을 한 줄에 하나씩 ------------------------------------
    ' 한글 찾아바꾸기는 줄바꿈 문자를 넣지 못한다. 그래서 서식 쪽에 줄을 미리
    ' 만들어 두고 줄마다 다른 태그를 심어, 항목을 한 줄에 하나씩 채운다.
    ' 서식이 가진 줄 수보다 항목이 적으면 뒤쪽은 빈 줄로 남는다.
    '
    ' 값이 없어도 **반드시 등록**해야 한다. 등록하지 않으면 RenderTemplate 이
    ' 그 태그를 바꾸지 않아 문서에 {{위험요인4}} 글자가 그대로 남는다.
    Dim i As Long
    For i = 1 To 5
        d("{{위험요인" & i & "}}") = NthItem(CStr(d("{{위험요인}}")), i)
        d("{{안전대책" & i & "}}") = NthItem(CStr(d("{{안전대책}}")), i)
    Next i
    For i = 1 To 3
        d("{{작업내용" & i & "}}") = NthItem(CStr(d("{{작업내용}}")), i)
        d("{{협의체안건" & i & "}}") = NthItem(CStr(d("{{협의체안건}}")), i)
        d("{{점검착안사항" & i & "}}") = NthItem(CStr(d("{{점검착안사항}}")), i)
        d("{{사용기계기구" & i & "}}") = NthItem(ContractTypeField(tname, DICT_기계기구), i)
        d("{{위험물질설비" & i & "}}") = NthItem(ContractTypeField(tname, DICT_위험물질), i)
        d("{{학교확인사항" & i & "}}") = NthItem(CStr(d("{{학교확인사항}}")), i)
        d("{{업체조치사항" & i & "}}") = NthItem(CStr(d("{{업체조치사항}}")), i)
        d("{{협의체합의문" & i & "}}") = NthItem(CStr(d("{{협의체합의문}}")), i)
        d("{{상시안건" & i & "}}") = NthItem(CStr(d("{{상시안건}}")), i)
    Next i

    Set BuildTagMap = d
End Function

' 앞의 것이 비면 뒤의 것을 쓴다.
Private Function OrElse(ByVal a As String, ByVal fallback As String) As String
    If Trim$(a) = "" Then
        OrElse = fallback
    Else
        OrElse = Trim$(a)
    End If
End Function

' " · " 로 이어 붙인 항목 목록에서 n 번째(1부터). 없으면 "".
'
' 사전 시트는 항목을 " · " 로 나누어 적고, MergeBlocks 는 사전 문구와
' ContentLibrary 공통 문구를 LINE_SEP("  ·  ") 로 잇는다. 둘 다 " · " 를
' 품고 있으므로 " · " 하나로 쪼갠 뒤 앞뒤 공백만 털면 둘 다 걸린다.
' '단부·개구부' 처럼 **공백 없는** 가운뎃점은 항목 안의 글자라 쪼개지지 않는다.
' 그래서 사전에 항목을 적을 때는 반드시 앞뒤에 공백을 둔 가운뎃점을 써야 한다.
Private Function NthItem(ByVal s As String, ByVal n As Long) As String
    If Trim$(s) = "" Then Exit Function

    Dim parts() As String, i As Long, k As Long, v As String
    parts = Split(s, " · ")
    For i = LBound(parts) To UBound(parts)
        v = Trim$(parts(i))
        If v <> "" Then
            k = k + 1
            If k = n Then
                NthItem = v
                Exit Function
            End If
        End If
    Next i
End Function

' 사전 문구를 앞에, 작업유형별 공통 문구를 뒤에 붙인다.
Private Function MergeBlocks(ByVal a As String, ByVal b As String) As String
    a = Trim$(a)
    b = Trim$(b)
    If a = "" Then
        MergeBlocks = b
    ElseIf b = "" Then
        MergeBlocks = a
    Else
        MergeBlocks = a & LINE_SEP & b
    End If
End Function

' 계약기간용 날짜. 여기에는 요일을 붙이지 않는다 ― {{계약기간}} 은
' '2026. 03. 01. ~ 2027. 02. 28.(365일간)' 처럼 두 날짜가 한 줄에 들어가는데
' 요일까지 넣으면 계획서 표 칸을 넘긴다.
Private Function FmtDate(ByVal v As Variant) As String
    If IsDate(v) Then FmtDate = Format$(CDate(v), "yyyy. mm. dd.")
End Function

' 날짜만 ― 2026.6.4.
'
' 0 을 채우지 않는다(2026.3.4. 이지 2026.03.04. 가 아니다). 순회점검표의
' '점검일' 칸이 17mm 밖에 안 되어 한 글자라도 짧아야 한다.
Public Function FmtInspectDate(ByVal d As Date) As String
    FmtInspectDate = Format$(d, "yyyy.m.d.")
End Function

' 요일만 ― (목)
'
' 순회점검표는 이것을 **날짜 아랫줄**에 따로 넣는다. 한 줄에 붙여 쓰면 좁은
' 칸에서 '2026.6.4.(' 까지만 들어가고 '목)' 만 다음 줄로 넘어가, 여는 괄호가
' 날짜에 매달린 꼴이 된다.
Public Function FmtDow(ByVal d As Date) As String
    FmtDow = "(" & Left$(WeekdayName2(Weekday(d)), 1) & ")"
End Function

' 날짜 + 요일 ― 2026.6.4.(목). 한 줄에 들어가는 자리에 쓴다.
'
' 주말·공휴일이라 다음 평일로 민 날에 예전에는 "* " 를 붙였다. 요일이 늘 함께
' 나오므로 점검요일과 다른 요일이 곧 '밀린 날' 이라는 표시가 된다. 별표는
' 뜻이 겹치기만 하고 칸만 잡아먹어 뺐다.
Public Function FmtDateDow(ByVal d As Date) As String
    FmtDateDow = FmtInspectDate(d) & FmtDow(d)
End Function

' 위와 같되 Variant 를 받는다. 날짜가 아니면 빈 문자열.
Private Function FmtDateW(ByVal v As Variant) As String
    If IsDate(v) Then FmtDateW = FmtDateDow(CDate(v))
End Function

' 주말이거나 공휴일이면 다음 평일로 미룬다. (NextWorkday ― modConfig)
Private Function FirstWeekday(ByVal v As Variant) As Variant
    If Not IsDate(v) Then Exit Function
    FirstWeekday = NextWorkday(CDate(v))
End Function

' ContentLibrary 시트를 통째로 읽어 둔 것. 셀을 하나씩 묻으면 계약마다 수천 번
' COM 을 오간다. 배열로 한 번에 떠서 그 위에서 돈다.
' (LibCache · BlockCache 선언은 모듈 맨 위, 비우는 것은 ResetExportCache)

Private Sub LoadLibCache()
    If Not IsEmpty(LibCache) Then Exit Sub

    Dim lib As Worksheet
    Set lib = ThisWorkbook.Worksheets(SH_LIB)

    Dim lastRow As Long
    lastRow = lib.Cells(lib.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    ' 1~4열만 쓴다(작업유형·구분·―·문구). .Value2 가 아니라 .Value 를 쓴다 ―
    ' 셀을 하나씩 읽던 예전 코드와 값이 한 글자도 달라지면 안 된다.
    LibCache = lib.Range(lib.Cells(1, 1), lib.Cells(lastRow, 4)).Value
End Sub

' ContentLibrary에서 해당 작업유형의 문구를 모아 한 덩어리로 만든다.
Private Function BuildContentBlock(ByVal types As String, ByVal category As String) As String
    On Error GoTo Fallback

    ' 같은 계약에서 위험요인·안전대책·상시안건으로 세 번 불린다. 답은 (types, category)
    ' 로 정해지므로 한 번 만든 것은 다시 쓴다.
    If BlockCache Is Nothing Then Set BlockCache = CreateObject("Scripting.Dictionary")
    Dim key As String
    key = types & vbTab & category
    If BlockCache.Exists(key) Then
        BuildContentBlock = BlockCache(key)
        Exit Function
    End If

    LoadLibCache

    Dim wanted As Object
    Set wanted = CreateObject("Scripting.Dictionary")
    wanted("공통") = True
    Dim t As Variant
    If types <> "" Then
        For Each t In Split(types, "|")
            wanted(CStr(t)) = True
        Next t
    End If

    Dim acc As String, r As Long
    For r = 2 To UBound(LibCache, 1)
        If wanted.Exists(Trim$(CStr(LibCache(r, 1) & ""))) Then
            If Trim$(CStr(LibCache(r, 2) & "")) = category Then
                If acc <> "" Then acc = acc & LINE_SEP
                acc = acc & CStr(LibCache(r, 4) & "")
            End If
        End If
    Next r

    BlockCache(key) = acc
    BuildContentBlock = acc
    Exit Function
Fallback:
    BuildContentBlock = ""
End Function

' 시트에서 떠 온 것을 모두 버린다. 버튼을 누를 때마다 부르므로, 사용자가 사전
' 시트를 고치고 곧바로 출력해도 옛 값이 나가지 않는다.
Public Sub ResetExportCache()
    LibCache = Empty
    Set BlockCache = Nothing
    ResetLookupCache                 ' modConfig 쪽 (사전 · 폴더 경로)
End Sub

' 서식 이름 → 그 서식에 실제로 든 태그 Dictionary.
'
' AllReplace 는 한 번마다 문서 전체를 훑는다. BuildTagMap 이 만드는 태그는 126개인데
' 서식 하나에 박혀 있는 것은 3~37개뿐이라, 전부 걸면 열에 아홉은 찾아도 안 나오는
' 것을 뒤지는 헛일이다. 서식 한 장에 100번씩 쌓이므로 그냥 넘길 양이 아니다.
'
' 목록은 modTemplateTags (gen_template_tags.py 가 서식 파일에서 뽑아 만든다).
' 비어 있으면 '모르는 서식' 이라는 뜻이고, 그때는 예전처럼 전부 바꾼다 ―
' 표가 낡아도 느려질 뿐 출력이 깨지지는 않는다.
' (TagSetCache 선언은 모듈 맨 위에 있다)

Private Function TemplateTagSet(ByVal tmplName As String) As Object
    If TagSetCache Is Nothing Then Set TagSetCache = CreateObject("Scripting.Dictionary")

    If Not TagSetCache.Exists(tmplName) Then
        Dim d As Object, parts() As String, i As Long, s As String
        Set d = CreateObject("Scripting.Dictionary")
        s = TemplateTagList(tmplName)
        If s <> "" Then
            parts = Split(s, "|")
            For i = LBound(parts) To UBound(parts)
                d.Item(parts(i)) = True
            Next i
        End If
        ' 개체를 사전에 넣을 때 'Set 사전(키) = 개체' 로 쓰면 한셀이 오류 450
        ' (인수의 개수나 속성 지정이 잘못되었습니다)을 낸다. 기본 속성을 거쳐
        ' 개체를 대입하는 것을 못 받는다 ― 문자열을 넣는 '사전(키) = 값' 은 받아
        ' 주므로 BuildTagMap 은 멀쩡하고 여기서만 걸렸다. Add·Item 으로 직접 부른다.
        TagSetCache.Add tmplName, d
    End If

    Set TemplateTagSet = TagSetCache.Item(tmplName)
End Function

' 템플릿을 열어 태그를 바꾸고 임시폴더에 저장한다. 반환값은 저장된 경로.
Public Function RenderTemplate(ByVal hwp As Object, ByVal tmplName As String, _
                               ByVal tags As Object, ByVal outStem As String) As String
    ' 넘어진 자리를 알려 주는 표시. 한셀에서는 엑셀과 다른 단계에서 멈추는데,
    ' 오류 번호만으로는 어디였는지 알 수가 없었다.
    Dim stage As String

    Dim src As String
    stage = "템플릿 폴더 찾기"
    src = TemplateFolder() & tmplName
    If Dir$(src) = "" Then
        Notify "서식 파일을 찾을 수 없습니다:" & vbLf & src & vbLf & vbLf & _
               "설정 시트의 템플릿 폴더 경로를 확인하세요.", vbExclamation, "서식 없음"
        Exit Function
    End If

    ' 임시폴더는 **열기 전에** 마련한다. 예전에는 저장 직전에 만들었는데, 그때
    ' 넘어지면 '서식 처리 중 오류' 로 둔갑해 서식 파일이 잘못된 것처럼 보였다.
    Dim tmpDir As String
    stage = "임시폴더 준비"
    On Error GoTo Failed
    tmpDir = TempFolder()

    ' 템플릿 원본은 **열지 않는다**. 임시폴더로 한 벌 복사해 그 사본을 연다.
    ' 한글의 '다른 이름으로 저장' 기본 폴더는 마지막으로 연 폴더라, 원본을 직접
    ' 열면 사용자가 실수로 템플릿을 덮어쓸 자리에 놓이게 된다.
    Dim work As String
    stage = "템플릿 복사"
    work = tmpDir & "~" & tmplName
    On Error Resume Next
    If Dir$(work) <> "" Then Kill work
    On Error GoTo Failed
    FileCopy src, work

    ' hwp.Open 의 반환값은 믿을 수 없다(문서가 제대로 열려도 False 가 돌아온다).
    ' 그래서 결과 파일이 실제로 만들어졌는지로 성공을 판단한다.
    stage = "서식 열기"
    hwp.Open work, "HWPX", "forceopen:true"

    ' 이 서식에 실제로 박혀 있는 태그만 바꾼다. (TemplateTagSet 주석 참고)
    Dim want As Object, replaceAllTags As Boolean
    stage = "태그 목록 읽기"
    Set want = TemplateTagSet(tmplName)
    replaceAllTags = (want.Count = 0)

    ' Keys 는 속성이 아니라 메서드다. 괄호 없이 쓰면 한셀이 걸고 넘어진다.
    Dim keyList As Variant, i As Long
    stage = "태그 치환"
    keyList = tags.Keys()
    For i = LBound(keyList) To UBound(keyList)
        If replaceAllTags Then
            ReplaceAll hwp, CStr(keyList(i)), CStr(tags(keyList(i)))
        ElseIf want.Exists(CStr(keyList(i))) Then
            ReplaceAll hwp, CStr(keyList(i)), CStr(tags(keyList(i)))
        End If
    Next i

    Dim out As String
    stage = "임시본 저장"
    ' 조각은 '~' 로 시작하는 임시본이다. 이어붙이고 나면 MergeIntoNewDoc 이 지운다.
    ' 사용자에게 남는 것은 이어붙인 한 파일뿐이다.
    out = tmpDir & "~" & SafeName(outStem) & ".hwpx"
    hwp.SaveAs out, "HWPX", ""
    stage = "임시본 닫기"
    hwp.XHwpDocuments.Close False           ' 템플릿 원본이 아니라 임시본을 닫는다

    On Error Resume Next
    Kill work                               ' 작업용 사본은 남기지 않는다
    On Error GoTo 0

    If Dir$(out) <> "" Then RenderTemplate = out
    Exit Function

Failed:
    ' Err 을 **가장 먼저** 읽는다. On Error 문은 Err 을 지우므로, 아래 정리 코드를
    ' 먼저 두면 메시지가 늘 빈칸으로 나온다("처리 중 오류: " 뒤가 비어 원인을
    ' 알 수 없었다). 어느 단계에서 넘어졌는지도 함께 남긴다.
    Dim errNo As Long, errMsg As String
    errNo = Err.Number
    errMsg = Err.Description
    If Trim$(errMsg) = "" Then errMsg = "(설명 없음)"

    ' 템플릿을 열어 둔 채로 두면 사용자가 실수로 덮어쓸 수 있다. 반드시 닫는다.
    On Error Resume Next
    hwp.XHwpDocuments.Close False
    On Error GoTo 0
    LastError = "서식 '" & tmplName & "' 처리 중 오류 " & errNo & ": " & errMsg & _
                " [단계: " & stage & "]"
End Function

Private Sub ReplaceAll(ByVal hwp As Object, ByVal findStr As String, ByVal replaceStr As String)
    On Error Resume Next
    hwp.HAction.GetDefault "AllReplace", hwp.HParameterSet.HFindReplace.HSet
    With hwp.HParameterSet.HFindReplace
        .FindString = findStr
        .ReplaceString = replaceStr
        .IgnoreMessage = 1
        .Direction = 0
        .MatchCase = 0
        .SeveralWords = 0
        .UseWildCards = 0
        .WholeWordOnly = 0
        .AutoSpell = 0
        .ReplaceMode = 1
    End With
    hwp.HAction.Execute "AllReplace", hwp.HParameterSet.HFindReplace.HSet
    On Error GoTo 0
End Sub

' 여러 임시본을 새 문서 하나로 이어붙인다.
' 서식과 서식 사이에는 반드시 쪽 나눔을 넣는다. 그래야 서식마다 새 쪽 맨 위에서
' 시작한다. (한 서식이 A4 를 다 못 채우면 다음 서식이 그 아래에 붙어 버린다.
' InsertFile 만으로는 쪽이 넘어가지 않는다 ― 실측으로 확인했다)
' 한 서식이 여러 쪽인 것(예: 계획서 상세 3쪽)은 그대로 이어져도 된다.
'
' 쪽을 넘기는 방법은 **딱 하나여야 한다.** 예전에는 쪽 나눔과 구역 나눔
' (InsertFile 의 KeepSection = 1)을 둘 다 걸었는데, 한글 2022 는 구역 나눔만으로도
' 쪽을 넘겨 버려서 쪽 나눔이 만든 쪽이 아무도 채우지 않은 **빈 쪽**으로 남았다
' (한글 2024 는 안 넘겨서 멀쩡했다). 지금은 InsertFile 에서 구역을 가져오지 않고
' (KeepSection = 0) 쪽 나눔 하나만 쓰므로 한글 버전을 타지 않는다.
'
' **오류를 밖으로 새어 나가게 두면 안 된다.** 한글이 도중에 죽으면 hwp 개체가
' 껍데기만 남는데, 그 뒤의 Open·SaveAs·PageCount 가 모두 오류를 내고 그것이
' 처리되지 않은 채 위로 올라간다. 엑셀은 보이지 않는 상태이므로 오류창을 띄울
' 자리가 없어 **VBA 편집기를 중단 모드로 띄운 채 멈춘다** ― 사용자에게는
' '작업창이 잠깐 떴다 사라지는' 것으로 보인다. 여기서 붙잡아 LastError 에 적는다.
Private Function MergeIntoNewDoc(ByVal hwp As Object, ByRef paths() As String, _
                                 ByVal outStem As String) As Boolean
    Dim stage As String
    On Error GoTo Failed

    stage = "첫 조각 열기"
    hwp.Open paths(LBound(paths)), "HWPX", "forceopen:true"

    Dim i As Long, failed As Long
    stage = "조각 이어붙이기"
    For i = LBound(paths) + 1 To UBound(paths)
        If Not BreakToNewPage(hwp) Then failed = failed + 1
        InsertFile hwp, paths(i)
    Next i

    stage = "문서 맨 앞으로"
    hwp.HAction.Run "MoveDocBegin"

    ' 점검용 기록 (GetLastMergeInfo 로 읽는다)
    LastMergedFiles = UBound(paths) - LBound(paths) + 1
    LastMergedPages = hwp.PageCount
    LastBreakFails = failed

    ' 이어붙인 결과를 한 파일로 저장한다. 조각(~로 시작하는 임시본)은 지운다.
    Dim out As String
    stage = "결과 저장"
    out = TempFolder() & SafeName(outStem) & ".hwpx"
    hwp.SaveAs out, "HWPX", ""

    On Error Resume Next
    For i = LBound(paths) To UBound(paths)
        Kill paths(i)
    Next i
    On Error GoTo 0

    MergeIntoNewDoc = True
    Exit Function

Failed:
    Dim errNo As Long, errMsg As String
    errNo = Err.Number
    errMsg = Err.Description
    If Trim$(errMsg) = "" Then errMsg = "(설명 없음)"
    LastError = "'" & outStem & "' 이어붙이는 중 오류 " & errNo & ": " & errMsg & _
                " [단계: " & stage & "]"

    ' 조각은 남겨 둔다. 무엇까지 만들어졌는지 볼 수 있어야 원인을 찾는다.
    On Error Resume Next
    hwp.XHwpDocuments.Close False
    On Error GoTo 0
End Function

' 다음 서식은 **반드시 새 쪽에서** 시작해야 한다. 쪽 나눔이 먹지 않으면 앞 서식이
' 쓰다 남긴 자리에 다음 서식 앞부분이 그대로 이어 붙는다(실제로 그런 일이 있었다).
' 문서 끝이 표 안이면 BreakPage 가 표만 건드리고 쪽은 그대로다.
' 그래서 쪽 수가 늘었는지 세어 보고, 안 늘었으면 문단을 하나 만들어 다시 넘긴다.
Private Function BreakToNewPage(ByVal hwp As Object) As Boolean
    Dim before As Long
    hwp.HAction.Run "MoveDocEnd"
    before = hwp.PageCount
    hwp.HAction.Run "BreakPage"
    If hwp.PageCount > before Then
        BreakToNewPage = True
        Exit Function
    End If

    hwp.HAction.Run "MoveDocEnd"
    hwp.HAction.Run "BreakPara"          ' 표 밖에 새 문단을 만들고 다시 넘긴다
    hwp.HAction.Run "BreakPage"
    BreakToNewPage = (hwp.PageCount > before)
End Function

' 구역(용지·여백)은 가져오지 않는다. 가져오면 구역 나눔이 하나 더 들어가고,
' 한글 2022 에서는 그것이 쪽을 한 번 더 넘겨 서식마다 빈 쪽이 하나씩 생긴다.
' 이어붙는 서식들은 어차피 용지·여백이 모두 같으므로(01~10 한 묶음, 11~15 한 묶음,
' 두 묶음이 섞여 붙는 일은 없다) 첫 문서의 구역을 따라도 달라지는 것이 없다.
' verify_templates.py 가 이 전제를 검사한다.
Private Sub InsertFile(ByVal hwp As Object, ByVal path As String)
    On Error Resume Next
    hwp.HAction.GetDefault "InsertFile", hwp.HParameterSet.HInsertFile.HSet
    With hwp.HParameterSet.HInsertFile
        .filename = path
        .KeepSection = 0        ' 구역은 가져오지 않는다 (위 주석 참고)
        .KeepCharshape = 1
        .KeepParashape = 1
        .KeepStyle = 1
    End With
    hwp.HAction.Execute "InsertFile", hwp.HParameterSet.HInsertFile.HSet
    On Error GoTo 0
End Sub

Private Sub ShowHwp(ByVal hwp As Object)
    On Error Resume Next
    hwp.XHwpWindows.Item(0).Visible = True
    On Error GoTo 0
End Sub

' 이어붙인 결과 파일 이름: 계약명_이름_계약시작일
'   1~8 묶음      → 2026학년도 청소미화 위탁 용역_중대재해예방_20260301
'   순회 12장     → 2026학년도 청소미화 위탁 용역_순회점검표_총12회차_20260301
' 계약을 여러 건 골라 한 문서로 묶었으면 "_외 N건" 을 덧붙인다.
Private Function OutName(ByVal inputRow As Long, ByVal label As String, _
                         Optional ByVal moreRows As Long = 0) As String
    Dim ws As Worksheet
    Set ws = WsInput()

    Dim stem As String
    stem = Trim$(CStr(ws.Cells(inputRow, CC_계약명).Value)) & "_" & label
    If moreRows > 0 Then stem = stem & "_외" & moreRows & "건"

    Dim s As Variant
    s = ws.Cells(inputRow, CC_시작일).Value
    If IsDate(s) Then stem = stem & "_" & Format$(CDate(s), "yyyymmdd")

    OutName = stem
End Function

' 템플릿 파일 이름에서 서식 이름만 뗀다 ("09_순회점검표.hwpx" → "순회점검표")
Private Function LabelOf(ByVal tmplName As String) As String
    Dim s As String
    s = tmplName
    If InStrRev(s, ".") > 0 Then s = Left$(s, InStrRev(s, ".") - 1)
    If InStr(s, "_") > 0 Then s = Mid$(s, InStr(s, "_") + 1)
    LabelOf = s
End Function

Private Function SafeName(ByVal s As String) As String
    Dim bad As Variant, c As Variant
    ' 가운데점(·)도 뺀다 ― 계약명에 자주 들어가는데 그대로 두면 저장이 막히는 일이 있다.
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "·")
    For Each c In bad
        s = Replace(s, c, "")
    Next c
    SafeName = s
End Function

'==============================================================================
' 한글 자동화 보안모듈 등록 (설정 시트 버튼)
'==============================================================================

' 한글은 외부 프로그램이 문서를 열려고 하면 경고창을 띄운다. 경고창을 없애려면
' 신뢰할 수 있는 경로를 알려주는 보안 모듈 DLL을 HKCU에 등록해야 한다.
' (관리자 권한은 필요 없다)
Public Sub RegisterHwpSecurityModule()
    Dim dllPath As String
    dllPath = ThisWorkbook.Path & "\form\_bin\FilePathCheckerModule.dll"

    If Dir$(dllPath) = "" Then
        MsgBox "보안 모듈 파일을 찾을 수 없습니다:" & vbLf & dllPath & vbLf & vbLf & _
               "form\_bin 폴더에 FilePathCheckerModule.dll 이 있어야 합니다.", _
               vbExclamation, "보안모듈 등록"
        Exit Sub
    End If

    On Error GoTo Failed
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.RegWrite "HKCU\SOFTWARE\HNC\HwpAutomation\Modules\FilePathCheckerModule", _
                dllPath, "REG_SZ"

    MsgBox "한글 자동화 보안모듈을 등록했습니다." & vbLf & _
           "한글이 열려 있었다면 한 번 닫았다 여세요.", vbInformation, "보안모듈 등록"
    Exit Sub
Failed:
    MsgBox "보안모듈 등록에 실패했습니다: " & Err.Description, vbCritical, "보안모듈 등록"
End Sub
