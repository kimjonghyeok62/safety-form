Attribute VB_Name = "modNotice"
Option Explicit

'==============================================================================
' 출력 안내 창
'
' 버튼을 누르면 한글이 뜨기까지 몇 초에서 몇십 초가 걸린다. 그동안 아무 것도
' 안 보이면 눌린 것인지 알 수 없다. 그래서 '무엇을 몇 장 만드는지' 를 먼저
' 띄워 두고, 다 만들면 스스로 닫는다.
'
' MsgBox 는 글자 크기를 정할 수 없어 UserForm(frmNotice)을 쓴다.
' frmNotice 는 build_workbook.py 가 만들어 넣는다.
'==============================================================================

' 안내창 글자 크기·여백 (50대 노안 기준. 이보다 키우면 줄이 길어져 오히려 읽기 나쁘다)
Public Const NOTICE_FONT_SIZE As Long = 12
Private Const LINE_H  As Long = 17          ' 12pt 한 줄 높이
Private Const PAD     As Long = 12
Private Const MAX_H   As Long = 640

' 예상 소요시간 어림값 (EtaText). 실측이 아니므로 체감이 다르면 여기만 고친다.
'
' 1~8번 서식은 태그가 많고(계획서 상세는 37개에 3쪽) 순회·합동은 한 장에 7~12개라
' 장당 드는 시간이 다르다. 그래서 장당 값을 둘로 나눠 둔다.
Private Const ETA_기본초     As Long = 5    ' 한글이 뜨기까지
Public Const ETA_장당초      As Long = 1    ' 1~8번 · 11~15번 서식 한 장
Public Const ETA_반복장당초  As Long = 1    ' 순회 · 합동 한 장
Private Const ETA_최소초     As Long = 5

' 계약이 이보다 많으면 계약마다 한 줄로 줄여 쓴다. 다 늘어놓으면 화면을 넘는다.
Private Const DETAIL_LIMIT As Long = 3

' 서식에 찍히는 날짜 표기 설명. 서식 안에는 범례를 넣을 자리가 없어 여기서 알린다.
' (modHwpExport.FmtInspectDate 가 붙이는 그 별표다)
Private Const STAR_LEGEND As String = _
    "날짜 뒤 괄호는 요일입니다. 앞에 * 가 붙은 날은 주말·공휴일이라 다음 평일로 미룬 날입니다."

Public Sub ShowNotice(ByVal title As String, ByVal body As String)
    If SilentMode Then Exit Sub
    On Error GoTo Fallback

    Dim lines As Long
    lines = UBound(Split(body, vbLf)) + 1

    frmNotice.Caption = title
    frmNotice.lblBody.Font.Size = NOTICE_FONT_SIZE
    frmNotice.lblBody.Caption = body
    frmNotice.lblBody.Height = lines * LINE_H + 4
    frmNotice.Height = PAD + frmNotice.lblBody.Height + PAD + 26
    If frmNotice.Height > MAX_H Then frmNotice.Height = MAX_H

    frmNotice.Show vbModeless
    frmNotice.Repaint
    DoEvents
    Exit Sub

Fallback:
    ' 폼이 말을 안 들어도 출력 자체는 되어야 한다. 안내를 포기하고 넘어간다.
    On Error Resume Next
End Sub

Public Sub HideNotice()
    On Error Resume Next
    Unload frmNotice
    On Error GoTo 0
End Sub

'==============================================================================
' 안내 문구 만들기
'==============================================================================

' 계약 한 건의 머리말 ― 계약명·기간·요일
Private Function ContractHead(ByVal inputRow As Long) As String
    Dim ws As Worksheet: Set ws = WsInput()
    Dim s As Variant, e As Variant, acc As String

    acc = "계약명: " & CStr(ws.Cells(inputRow, CC_계약명).Value)

    s = ws.Cells(inputRow, CC_시작일).Value
    e = ws.Cells(inputRow, CC_종료일).Value
    If IsDate(s) And IsDate(e) Then
        acc = acc & vbLf & "계약기간: " & Format$(s, "yyyy-mm-dd") & " ~ " & _
              Format$(e, "yyyy-mm-dd") & "  (" & ContractDays(inputRow) & "일)"
    End If

    ContractHead = acc
End Function

' 1번 계획서는 계약 성격에 따라 셋 중 하나가 나간다. 어느 것인지 알려 준다.
' (추천 안내창도 이것을 쓰므로 Public 이다)
Public Function PlanKind(ByVal inputRow As Long) As String
    Dim f As String
    f = ResolvePlanTemplate(inputRow)
    If InStr(f, "상세") > 0 Then
        PlanKind = "상세"
    ElseIf InStr(f, "공사") > 0 Then
        PlanKind = "공사"
    Else
        PlanKind = "용역"
    End If
End Function

' 서식 번호는 출력관리 화면의 번호다(파일 이름 앞자리와 다르다 ―
' modHwpExport.TemplateFileFor 주석 참고). 8번은 자기 버튼으로 따로 나가므로
' 여기 1~7 묶음 안내에는 뜨지 않지만, 이름표는 맞춰 둔다.
Private Function FormTitle(ByVal formNo As Long, ByVal inputRow As Long) As String
    Select Case formNo
        Case 1: FormTitle = "안전보건관리계획서 (" & PlanKind(inputRow) & ")"
        Case 2: FormTitle = "수급인 안전보건 수준평가표"
        Case 3: FormTitle = "공사(용역) 안전보건 점검표"
        Case 4: FormTitle = "점검표 (고소작업)"
        Case 5: FormTitle = "점검표 (화기작업)"
        Case 6: FormTitle = "점검표 (밀폐공간)"
        Case 7: FormTitle = "점검표 (전기작업)"
        Case 8: FormTitle = "안전보건협의체 회의록"
        Case Else: FormTitle = ""
    End Select
End Function

' 그 서식이 왜 필요한지 ― 판정 기준. 숫자는 설정 시트에서 읽는다.
Private Function FormRule(ByVal formNo As Long) As String
    Select Case formNo
        Case 1: FormRule = CriteriaPlan() & "일 이상"
        Case 2: FormRule = CriteriaLevel() & "일 이상"
        Case 3: FormRule = "모든 계약"
        Case 4: FormRule = "고소작업 시"
        Case 5: FormRule = "화기작업 시"
        Case 6: FormRule = "밀폐공간 시"
        Case 7: FormRule = "전기작업 시"
        Case 8: FormRule = CriteriaCouncil() & "일 이상 · 월 1회"
        Case Else: FormRule = ""
    End Select
End Function

' 그 계약에 ○ 표시된 1~7 서식 번호들
Private Function MarkedForms(ByVal inputRow As Long) As Variant
    Dim wsOut As Worksheet: Set wsOut = WsOutput()
    Dim oRow As Long: oRow = OutRow(inputRow)

    Dim acc() As Long, n As Long, formNo As Long
    ReDim acc(0 To FORM_MERGE_LAST - 1)
    For formNo = 1 To FORM_MERGE_LAST
        If IsMarked(wsOut.Cells(oRow, FormCol(formNo)).Value) Then
            acc(n) = formNo
            n = n + 1
        End If
    Next formNo

    If n = 0 Then
        MarkedForms = Array()
    Else
        ReDim Preserve acc(0 To n - 1)
        MarkedForms = acc
    End If
End Function

' [1~7번 출력] 안내문
Public Function NoticeText17(ByVal rows As Variant) As String
    Dim acc As String, i As Long, inputRow As Long
    Dim forms As Variant, j As Long, brief As Boolean
    Dim totalForms As Long

    brief = (UBound(rows) - LBound(rows) + 1 > DETAIL_LIMIT)

    For i = LBound(rows) To UBound(rows)
        inputRow = CLng(rows(i))
        forms = MarkedForms(inputRow)
        totalForms = totalForms + (UBound(forms) + 1)

        If acc <> "" Then acc = acc & vbLf & vbLf
        If brief Then
            acc = acc & " · " & CStr(WsInput().Cells(inputRow, CC_계약명).Value) & _
                  "  →  " & (UBound(forms) + 1) & "종"
        Else
            acc = acc & ContractHead(inputRow) & vbLf & vbLf & "출력물"
            For j = LBound(forms) To UBound(forms)
                acc = acc & vbLf & "  " & forms(j) & ". " & _
                      FormTitle(CLng(forms(j)), inputRow) & _
                      "  ―  " & FormRule(CLng(forms(j)))
            Next j
            acc = acc & vbLf & "  → 모두 " & (UBound(forms) + 1) & "종"
        End If
    Next i

    If brief Then
        acc = "계약 " & (UBound(rows) - LBound(rows) + 1) & "건을 출력합니다." & _
              vbLf & vbLf & acc
    End If

    ' 고른 계약을 전부 한 문서에 이어 붙인다. 계약과 계약 사이에는 쪽 나눔이 들어간다.
    If UBound(rows) - LBound(rows) + 1 > 1 Then
        acc = acc & vbLf & vbLf & "고른 계약 " & (UBound(rows) - LBound(rows) + 1) & _
              "건을 한 문서로 이어 출력합니다. (계약 사이는 쪽이 넘어갑니다)"
    Else
        acc = acc & vbLf & vbLf & "○ 표시된 서식을 한 문서로 이어 출력합니다."
    End If

    ' 계획서에는 파란 예시 글이 남아 있다. 그 자리는 사용자가 손봐야 한다.
    NoticeText17 = acc & vbLf & vbLf & _
                   "(파란색 부분은 수정해주세요)" & vbLf & vbLf & _
                   "한글에서 문서가 열릴 때까지 잠시 기다려 주세요. " & _
                   EtaText(totalForms, ETA_장당초)
End Function

' 서식 장수로 짐작한 대기 시간. 한글을 띄우는 데 드는 고정 시간에
' 서식 한 장당 드는 시간을 더한다. 실측이 아니라 어림이므로 '약'을 붙여 보여 준다.
' 체감이 다르면 모듈 맨 위 ETA_ 상수 네 개만 고치면 된다.
'
' perPage 를 안 주면 1~8번 기준(ETA_장당초)으로 잡는다.
Public Function EtaText(ByVal pageCount As Long, _
                        Optional ByVal perPage As Long = 0) As String
    If perPage <= 0 Then perPage = ETA_장당초

    Dim sec As Long
    sec = ETA_기본초 + pageCount * perPage
    If sec < ETA_최소초 Then sec = ETA_최소초

    If sec < 60 Then
        EtaText = "(예상 소요시간: 약 " & sec & "초)"
    ElseIf sec Mod 60 = 0 Then
        EtaText = "(예상 소요시간: 약 " & (sec \ 60) & "분)"
    Else
        EtaText = "(예상 소요시간: 약 " & (sec \ 60) & "분 " & (sec Mod 60) & "초)"
    End If
End Function

' [9. 순회점검표] / [10. 합동점검표] 안내문
'
' quarterlyOnly = True 면 합동(분기 1회, 90일 구간의 마지막 점검일),
' False 면 순회(매주, 달마다 1장).
Public Function NoticeTextRepeat(ByVal rows As Variant, ByVal quarterlyOnly As Boolean, _
                                 ByVal minDays As Long) As String
    Dim ws As Worksheet: Set ws = WsInput()
    Dim acc As String, i As Long, inputRow As Long
    Dim dates As Variant, wd As String, days As Long
    Dim totalPages As Long

    For i = LBound(rows) To UBound(rows)
        inputRow = CLng(rows(i))
        days = ContractDays(inputRow)
        wd = Trim$(CStr(ws.Cells(inputRow, CC_점검요일).Value))

        If acc <> "" Then acc = acc & vbLf & vbLf
        acc = acc & ContractHead(inputRow)

        If days < minDays Then
            acc = acc & vbLf & "  → 계약기간이 " & minDays & "일 미만이라 만들지 않습니다."
        Else
            dates = InspectionDates(inputRow, quarterlyOnly)
            If UBound(dates) < 0 Then
                acc = acc & vbLf & "  → 계약기간 안에 " & wd & "이 없어 만들지 않습니다."
            ElseIf quarterlyOnly Then
                acc = acc & vbLf & "점검요일: 90일 구간의 마지막 " & wd
                acc = acc & vbLf & vbLf & _
                      "  · 분기 1회 · 모두 " & (UBound(dates) + 1) & "회차" & vbLf & _
                      "  · 첫 점검일 " & Format$(dates(0), "yyyy-mm-dd") & _
                      " · 마지막 " & Format$(dates(UBound(dates)), "yyyy-mm-dd") & vbLf & _
                      "  → 합동점검표 " & (UBound(dates) + 1) & "장을 한 문서로 이어 출력합니다."
                totalPages = totalPages + (UBound(dates) + 1)
            Else
                Dim months As Collection
                Set months = PatrolMonths(inputRow)
                acc = acc & vbLf & "점검요일: 매주 " & wd
                acc = acc & vbLf & vbLf & _
                      "  · 점검일 모두 " & (UBound(dates) + 1) & "회" & vbLf & _
                      "  · 첫 점검일 " & Format$(dates(0), "yyyy-mm-dd") & _
                      " · 마지막 " & Format$(dates(UBound(dates)), "yyyy-mm-dd") & vbLf & _
                      "  · 한 장이 한 달치 (1주~5주 칸에 나눠 넣습니다)" & vbLf & _
                      "  → 순회점검표 " & months.Count & "장을 한 문서로 이어 출력합니다."
                totalPages = totalPages + months.Count
            End If
        End If
    Next i

    acc = acc & vbLf & vbLf & STAR_LEGEND

    ' 만들 게 하나도 없으면 기다릴 일도 없으므로 시간을 붙이지 않는다.
    acc = acc & vbLf & vbLf & "한글에서 문서가 열릴 때까지 잠시 기다려 주세요."
    If totalPages > 0 Then acc = acc & " " & EtaText(totalPages, ETA_반복장당초)

    NoticeTextRepeat = acc
End Function

'==============================================================================
' [8. 안전보건협의체 회의록] 안내문 ― 월 1회, 그 달의 마지막 순회점검일
'
' 순회/합동과 회차 잡는 규칙이 달라(달마다 마지막 하나) NoticeTextRepeat 의
' 분기를 늘리지 않고 따로 둔다. 그쪽은 순회·합동 두 갈래만 맡는다.
'==============================================================================

Public Function NoticeTextCouncil(ByVal rows As Variant) As String
    Dim ws As Worksheet: Set ws = WsInput()
    Dim acc As String, i As Long, inputRow As Long
    Dim dates As Variant, wd As String, days As Long, minDays As Long
    Dim totalPages As Long

    minDays = CriteriaCouncil()

    For i = LBound(rows) To UBound(rows)
        inputRow = CLng(rows(i))
        days = ContractDays(inputRow)
        wd = Trim$(CStr(ws.Cells(inputRow, CC_점검요일).Value))

        If acc <> "" Then acc = acc & vbLf & vbLf
        acc = acc & ContractHead(inputRow)

        If days < minDays Then
            acc = acc & vbLf & "  → 계약기간이 " & minDays & "일 미만이라 만들지 않습니다."
        Else
            dates = CouncilDates(inputRow)
            If UBound(dates) < 0 Then
                acc = acc & vbLf & "  → 계약기간 안에 " & wd & "이 없어 만들지 않습니다."
            Else
                acc = acc & vbLf & "회의일: 달마다 마지막 " & wd
                acc = acc & vbLf & vbLf & _
                      "  · 월 1회 · 모두 " & (UBound(dates) + 1) & "회차" & vbLf & _
                      "  · 첫 회의일 " & Format$(dates(0), "yyyy-mm-dd") & _
                      " · 마지막 " & Format$(dates(UBound(dates)), "yyyy-mm-dd") & vbLf & _
                      "  · 회차마다 안건과 의결사항이 조금씩 다릅니다" & vbLf & _
                      "  → 회의록 " & (UBound(dates) + 1) & "장을 한 문서로 이어 출력합니다." & vbLf & _
                      "    (회차마다 새 쪽 맨 위에서 시작합니다)"
                totalPages = totalPages + (UBound(dates) + 1)
            End If
        End If
    Next i

    acc = acc & vbLf & vbLf & STAR_LEGEND
    acc = acc & vbLf & vbLf & "한글에서 문서가 열릴 때까지 잠시 기다려 주세요."
    If totalPages > 0 Then acc = acc & " " & EtaText(totalPages, ETA_반복장당초)

    NoticeTextCouncil = acc
End Function
