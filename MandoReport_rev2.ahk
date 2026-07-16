; 사용기한 체크 (2027년 12월 31일 23:59:59까지 허용)
ExpireDate := "20271231235959"  ; 만료일자

if (A_Now > ExpireDate)
{
    MsgBox("관리자에게 문의하세요.`n프로그램을 종료합니다.", "알림", 48)
    ExitApp()
}

; ======================== CMM 엑셀 저장 설정 ========================
CFG_PartNoX := 460       ; 품번 셀 좌표 (헤더 보임 상태)
CFG_PartNoY := 380
CFG_ItemX := 394         ; 검사항목 컬럼 x좌표 (그리드)
CFG_DetailX := 442       ; 세부내역 컬럼 x좌표 (그리드)
CFG_SeqX := 305          ; 순번 컬럼 x좌표 (그리드) - 행 중복처리 방지용
CFG_TemplatePath := A_ScriptDir . "\Master.xlsx"   ; 엑셀 양식 파일
CFG_SaveFolder := A_ScriptDir . "\CMM_Result"      ; 결과 저장 폴더
; ===================================================================

; CMM 수집 데이터 (F1 실행마다 초기화)
CMM_Records := []
PartNo := ""

F2::ExitApp()

; F3: 현재 마우스 좌표 확인 (좌표 세팅용)
F3::
{
    MouseGetPos(&mx, &my)
    A_Clipboard := mx . ", " . my
    ToolTip("좌표: " . mx . ", " . my . " (클립보드에 복사됨)")
    SetTimer(() => ToolTip(), -2000)
}

; 계측기가 삼차원 측정기(CMM) 계열인지 판별
IsCMM(Equipment)
{
    if (!Equipment)
        return false
    return (InStr(Equipment, "삼차원") || InStr(Equipment, "CMM") || SubStr(Equipment, 1, 2) = "삼차")
}

Ave(Lower, Target, Upper, Equipment) ; 랜덤값 계산식
{
    ; 원본 문자열 값 보존 (0.0과 0 구분을 위해)
    OriginalUpper := Upper

    ; 입력값 유효성 검사
    try {
        Lower := Float(Lower)
        Target := Float(Target)
        Upper := Float(Upper)
    }
    catch {
        ; 숫자가 아닌 경우 기본값 반환
        return 0
    }

    ; 문자열 앞 2글자 추출 (null 체크 추가)
    if (!Equipment || StrLen(Equipment) < 2) {
        prefix := ""
    } else {
        prefix := SubStr(Equipment, 1, 2)
    }

    ; 새로운 룰 적용 (우선순위대로)
    ; 1. upper값이 정확히 "0.0"이고 equipment가 로크일 때
    if (OriginalUpper = "0.0" && prefix = "로크") {
        Rlower := 78
        Rupper := 82
        ;MsgBox("1번입니다")
    }
    ; 2. upper값이 정확히 "0.0"이고 equipment가 비중일 때
    else if (OriginalUpper = "0.0" && prefix = "비중") {
        Rlower := 7.02
        Rupper := 7.08
        ;MsgBox("2번입니다")
    }
    ; 3. lower값과 target값이 동일하면
    else if (Lower = Target) {
        Rlower := Upper * 0.1  ; 상한*10%
        Rupper := Upper * 0.3  ; 상한*30%
        ;MsgBox("3번입니다")
    }
    ; 4. target-lower값이 0.007이면
    else if (Abs(Target - Lower - 0.007) < 0.0001) {
        Rlower := Target - 0.002
        Rupper := Target + 0.002
        ;MsgBox("4번입니다")
    }
    ; 5. target-lower값이 0.01이면
    else if (Abs(Target - Lower - 0.01) < 0.0001) {
        Rlower := Target - 0.002
        Rupper := Target + 0.002
        ;MsgBox("5번입니다")
    }
    ; 6. target-lower값이 0.015이면
    else if (Abs(Target - Lower - 0.015) < 0.0001) {
        Rlower := Target - 0.003
        Rupper := Target + 0.003
        ;MsgBox("6번입니다")
    }
    ; 9. upper값이 정확히 "0.0"이고 equipment가 utm계열일 때
    else if (OriginalUpper = "0.0" && prefix = "전기" && Target < 100) {
        Rlower := 60
        Rupper := 80
        ;MsgBox("9번입니다")
    }
    ; 10. upper값이 정확히 "0.0"이고 equipment가 utm계열일 때
    else if (OriginalUpper = "0.0" && prefix = "전기" && Target > 100) {
        Rlower := Target - (Target - Lower) / 4
        Rupper := Target + (Target - Lower) / 4
        ;MsgBox("10번입니다")
    }

    ; 7. target-lower값이 0.05이상이면 (단, 로크는 제외)
    else if ((Target - Lower) >= 0.05 && prefix != "로크") {
        Rlower := Target - 0.03
        Rupper := Target + 0.03
        ;MsgBox("7번입니다")
    }
    ; 8. equipment가 게이지일 경우
    else if (prefix = "게이") {
        Rlower := Target - 0.01
        Rupper := Target + 0.01
        ;MsgBox("8번입니다")
    }
    ; 11. 기타 해당없는 경우 (기존 룰)
    else {
        Rlower := Target - (Target - Lower) / 4
        Rupper := Target + (Target - Lower) / 4
        ;MsgBox("11번입니다")
    }

    ; 랜덤 값 생성
    if (Rlower >= Rupper) {
        ; Rlower가 Rupper보다 크거나 같으면 Target 값 반환
        g := Target
    } else {
        g := Random(Rlower, Rupper)
    }

    ; 계측기별 소수점 처리
    if (prefix = "게이"
    || prefix = "버니"
    || prefix = "조도"
    || prefix = "외측"
    || prefix = "내측"
    || prefix = "비중"
    || prefix = "밀도")
    {
        h := Round(g, 2)
    }
    else if (prefix = "하이"
    || prefix = "인디"
    || prefix = "형상"
    || prefix = "전용"
    || prefix = "공구"
    || prefix = "삼차")
    {
        h := Round(g, 3)
    }
    else if (prefix = "경도"
    || prefix = "로크"
    || prefix = "파손")
    {
        h := Round(g, 1)
    }
    else if (prefix = "파괴"
    || prefix = "UT"
    || prefix = "전기")
    {
        h := Round(g, 0)
    }
    else
    {
        h := Round(g, 2)
    }

    return h
}

; 클립보드 복사를 재시도하는 함수
; 매개변수: x, y (좌표), clicks (클릭횟수, 기본값=2), maxRetries (최대재시도횟수, 기본값=3), timeout (각 시도별 타임아웃, 기본값=1초)
SafeCopy(x, y, clicks := 2, maxRetries := 3, timeout := 2) {
    retryCount := 0

    Loop {
        retryCount++

        ; 클립보드 초기화
        A_Clipboard := ""

        ; 클릭 및 복사
        Click(x, y, clicks)
        Sleep(50)
        Send("^c")

        ; 클립보드에 데이터가 들어올 때까지 대기
        if (ClipWait(timeout)) {
            ; 복사 성공 - 데이터가 비어있지 않은지 확인
            if (Trim(A_Clipboard) != "") {
                return A_Clipboard  ; 성공적으로 복사된 데이터 반환
            }
        }

        ; 최대 재시도 횟수에 도달했는지 확인
        if (retryCount >= maxRetries) {
            MsgBox("클립보드 복사 실패 - 좌표(" . x . ", " . y . ")에서 " . maxRetries . "번 시도 후 실패")
            return ""  ; 실패시 빈 문자열 반환
        }

        ; 재시도 전 잠시 대기
        Sleep(100)
    }
}

; 클립보드 복사를 재시도하는 함수
; 매개변수: x, y (좌표), clicks (클릭횟수, 기본값=3), maxRetries (최대재시도횟수, 기본값=3), timeout (각 시도별 타임아웃, 기본값=1초)
; returnZeroIfNotNumber: true이면 숫자가 아닐 때 0.0 반환, false이면 원본 반환
SafeCopy_1(x, y, clicks := 3, maxRetries := 3, timeout := 1, returnZeroIfNotNumber := false) {
    retryCount := 0

    Loop {
        retryCount++

        ; 클립보드 초기화
        A_Clipboard := ""

        ; 클릭 및 복사
        Click(x, y, clicks)
        Sleep(50)
        Send("^c")

        ; 클립보드에 데이터가 들어올 때까지 대기
        if (ClipWait(timeout)) {
            ; 복사 성공 - 데이터가 비어있지 않은지 확인
            if (Trim(A_Clipboard) != "") {
                ; 숫자가 아닌 경우 0.0 반환 옵션 체크
                if (returnZeroIfNotNumber) {
                    try {
                        ; 숫자로 변환 시도
                        Float(A_Clipboard)
                        return A_Clipboard  ; 숫자면 원본 반환
                    }
                    catch {
                        return "0.0"  ; 숫자가 아니면 0.0 반환
                    }
                } else {
                    return A_Clipboard  ; 원본 데이터 반환
                }
            }
        }

        ; 최대 재시도 횟수에 도달했는지 확인
        if (retryCount >= maxRetries) {
            MsgBox("클립보드 복사 실패 - 좌표(" . x . ", " . y . ")에서 " . maxRetries . "번 시도 후 실패")
            if (returnZeroIfNotNumber) {
                return "0.0"  ; 실패시에도 0.0 반환
            } else {
                return ""  ; 실패시 빈 문자열 반환
            }
        }

        ; 재시도 전 잠시 대기
        Sleep(100)
    }
}

; CMM 행 스캔: 순번/검사항목/세부내역을 읽고, Tab 기준점(계측기 셀)을 복원한다
; 반환값: Characteristic 문자열 (예: "180_대칭도", "100_Width_거리")
ScanCMMInfo(rowY)
{
    global CFG_SeqX, CFG_ItemX, CFG_DetailX

    Seq := Trim(SafeCopy(CFG_SeqX, rowY))
    ItemName := Trim(SafeCopy(CFG_ItemX, rowY))
    Detail := Trim(SafeCopy(CFG_DetailX, rowY))

    ; Tab 이동 기준점 복원 (계측기 셀 재클릭)
    Click(1235, rowY, 2)
    Sleep(50)

    if (ItemName != "" && Detail != "" && ItemName != Detail)
        name := ItemName . "_" . Detail
    else if (Detail != "")
        name := Detail
    else
        name := ItemName

    if (Seq != "")
        name := Seq . "_" . name

    return name
}

; 작업 종료 처리: CMM 수집분이 있으면 엑셀 저장 후 완료 메시지 표시
FinishReport(prefix)
{
    global CMM_Records, PartNo

    msg := prefix
    if (CMM_Records.Length > 0)
    {
        try {
            savedFile := SaveCMMToExcel(PartNo, CMM_Records)
            msg .= "`n`nCMM 측정값 " . CMM_Records.Length . "건을 저장했습니다:`n" . savedFile
        }
        catch as e {
            msg .= "`n`n단, CMM 엑셀 저장에 실패했습니다:`n" . e.Message
        }
    }
    MsgBox(msg)
}

; 수집된 CMM 측정값을 Master.xlsx 양식에 기록하여 저장
; 반환값: 저장된 파일 전체 경로
SaveCMMToExcel(partNo, records)
{
    global CFG_TemplatePath, CFG_SaveFolder

    if (!FileExist(CFG_TemplatePath))
        throw Error("엑셀 양식 파일을 찾을 수 없습니다:`n" . CFG_TemplatePath)

    if (!DirExist(CFG_SaveFolder))
        DirCreate(CFG_SaveFolder)

    ; 파일명에 쓸 수 없는 문자 제거
    safePartNo := RegExReplace(partNo, '[\\/:*?"<>|\r\n]', "")
    if (safePartNo = "")
        safePartNo := "NONAME"

    xl := ComObject("Excel.Application")
    xl.Visible := false
    xl.DisplayAlerts := false

    try {
        wb := xl.Workbooks.Open(CFG_TemplatePath, , true)  ; 읽기전용으로 열어 양식 보호
        ws := wb.Sheets(1)

        ; 헤더 채우기
        ws.Range("A5").Value := partNo  ; Drawing No. 자리에 품번
        ws.Range("C3").Value := FormatTime(A_Now, "yyyy-MM-dd")  ; Date (=NOW() 수식을 고정값으로)
        ws.Range("C5").Value := FormatTime(A_Now, "HH:mm:ss")    ; Time

        ; 10행부터 항목별 기록
        row := 10
        for rec in records {
            ws.Cells(row, 1).Value := rec.name             ; A: Characteristic
            ws.Cells(row, 2).Value := rec.target           ; B: Nominal
            ws.Cells(row, 3).Value := Round(rec.upper - rec.target, 4)  ; C: Upper Tol
            ws.Cells(row, 4).Value := Round(rec.lower - rec.target, 4)  ; D: Lower Tol
            col := 6                                       ; F열부터 #1~#5
            for v in rec.values {
                ws.Cells(row, col).Value := v
                col++
            }
            row++
        }

        fileName := CFG_SaveFolder . "\" . safePartNo . "_" . FormatTime(A_Now, "yyyyMMdd_HHmmss") . ".xlsx"
        wb.SaveAs(fileName, 51)  ; 51 = xlsx
        wb.Close(false)
        return fileName
    }
    finally {
        xl.Quit()
    }
}



F1::
{
    global CMM_Records, PartNo, CFG_PartNoX, CFG_PartNoY

    ; ===== 품번 확인 (헤더가 보이는 상태에서 F1을 누른다) =====
    CMM_Records := []
    PartNo := Trim(SafeCopy(CFG_PartNoX, CFG_PartNoY))

    answer := MsgBox("품번: [" . PartNo . "]`n`n'숨김' 버튼을 눌러 목록 화면으로 전환한 뒤 [확인]을 누르세요.", "품번 확인", "OKCancel")
    if (answer = "Cancel")
        return

    Click(872, 274)
    Sleep(80)
    Click(872, 274)
    Sleep(80)
    Click(1672, 274)
    Sleep(80)

    Data := SafeCopy(1664, 830) ; 건수 확인
    ;MsgBox(Data)

    DataCount := Data + 1

    x1 := 670  ; 구분 조회좌표 (정량 유무)
    y := 387
    y1 := 767
    Count := 0
    ; Y축 스크롤바를 위아래로 살짝 움직인 뒤 맨 위에 놓는다 (이후 한 칸씩 스크롤되도록 보정)
    ; 웹 스크롤바가 드래그를 인식하도록 느린 속도로 실행
    SetDefaultMouseSpeed(50)
    MouseClickDrag("Left", 1690, 460, 1690, 780)   ; 아래로 (최대한 천천히)
    Sleep(500)
    SetDefaultMouseSpeed(30)
    MouseClickDrag("Left", 1690, 780, 1690, 400)   ; 위로 초과 드래그하여 맨 위에 고정
    Sleep(500)
    SetDefaultMouseSpeed(2)
    ;MsgBox(DataCount)

    Loop
    {
        Count := Count + 1
        ;MsgBox("번호는 " . Count . " 째입니다.")

        if (Count = DataCount)
        {
            FinishReport("완료하였습니다.")
            return
        }
        else if (Count <= 15)
        {
            DataValue := SafeCopy(x1, y)
            ;MsgBox(DataValue)

            ; 유효성 검사 (재확인 로직)
            if (DataValue != "정량" && DataValue != "정성" && DataValue != "첨부") {
                Sleep(200)
                DataValue := SafeCopy(x1, y)
                MsgBox("재시도 결과: [" . DataValue . "]")

                if (DataValue != "정량" && DataValue != "정성" && DataValue != "첨부") {
                    MsgBox("구분값 인식 실패: [" . DataValue . "]`n수동으로 확인 후 계속하시겠습니까?")
                }
            }


            if (DataValue = "정량")
            {
                Lower := SafeCopy_1(860, y, 3, 3, 1, true) ; lower 확인
                Target := SafeCopy_1(945, y, 3, 3, 1, true) ; target 확인
                Upper := SafeCopy_1(992, y, 3, 3, 1, true) ; Upper 확인
                c := SafeCopy(1210, y) ; sample수 확인
                Equipment := SafeCopy(1235, y) ; 계측기 확인

                ; CMM이면 검사항목/세부내역 추가 스캔 (기준점은 함수 내에서 복원)
                isCMMRow := IsCMM(Equipment)
                if (isCMMRow)
                {
                    CharName := ScanCMMInfo(y)
                    cmmValues := []
                }

                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Loop c
                {
                    Sleep(100)
                    result := Ave(Lower, Target, Upper, Equipment)
                    Send(result)
                    if (isCMMRow)
                        cmmValues.Push(result)
                    Sleep(100)
                    Send("{Tab}")
                }

                ; CMM 항목이면 수집 목록에 추가
                if (isCMMRow && cmmValues.Length > 0)
                {
                    fLower := 0, fTarget := 0, fUpper := 0
                    try fLower := Float(Lower)
                    try fTarget := Float(Target)
                    try fUpper := Float(Upper)
                    CMM_Records.Push({name: CharName, lower: fLower, target: fTarget, upper: fUpper, values: cmmValues})
                }

                ; 정량이고 샘플수가 5일 때만 좌우 스크롤
                if (c >= 5) {
                    Sleep(50)
                    SetDefaultMouseSpeed(15)
                    MouseClickDrag("Left", 1170, 809, 700, 809)
                    Sleep(300)  ; 웹 테이블 정렬 완료까지 대기
                    MouseClickDrag("Left", 1170, 809, 700, 809)
                    Sleep(300)  ; 웹 테이블 정렬 완료까지 대기
                    SetDefaultMouseSpeed(2)
                }
            }
            else if (DataValue = "정성" || DataValue = "첨부")
            {
                c := SafeCopy(1210, y) ; sample수 확인
                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Loop 1
                {
                    Sleep(100)
                    Send("OK")
                    Sleep(100)
                    Send("{Tab}")
                }
                ; 정성/첨부는 스크롤 불필요
            }
            else
            {
                MsgBox("오류. 화면을닫고 재시작하세요")
                return
            }

            y := y + 27  ; 스크롤과 관계없이 다음 행으로 이동
        }
        else if (Count > 15)
        {
            ; 첫 진입 시 y1 위치(=마지막으로 처리한 15번째 행)의 순번을 기준값으로 저장
            if (Count = 16)
                LastSeq := Trim(SafeCopy(CFG_SeqX, y1))

            ; 순번이 바뀔 때까지 스크롤 (같은 행 중복 처리 방지)
            scrollTry := 0
            Seq := ""
            Loop
            {
                SetDefaultMouseSpeed(15)
                MouseClickDrag("Left", 1680, 550, 1680, 560) ; Y축 스크롤바
                Sleep(300)
                SetDefaultMouseSpeed(2)
                Sleep(50)
                Seq := Trim(SafeCopy(CFG_SeqX, y1))

                if (Seq != "" && Seq != LastSeq)
                    break  ; 새로운 행 확인됨

                scrollTry++
                if (scrollTry >= 3)
                {
                    ; 3회 스크롤해도 순번이 그대로면 그리드 끝에 도달한 것으로 판단
                    FinishReport("더 이상 새로운 행이 없어 종료합니다.`n(처리 " . (Count - 1) . "건 / 건수 " . Data . "건)")
                    return
                }
            }
            LastSeq := Seq

            DataValue := SafeCopy(x1, y1)
            ;MsgBox(DataValue)

            ; 유효성 검사 (재확인 로직)
            if (DataValue != "정량" && DataValue != "정성" && DataValue != "첨부") {
                Sleep(200)
                DataValue := SafeCopy(x1, y1)
                MsgBox("재시도 결과: [" . DataValue . "]")

                if (DataValue != "정량" && DataValue != "정성" && DataValue != "첨부") {
                    MsgBox("구분값 인식 실패: [" . DataValue . "]`n수동으로 확인 후 계속하시겠습니까?")
                }
            }


            if (DataValue = "정량")
            {
                Lower := SafeCopy_1(860, y1, 3, 3, 1, true) ; lower 확인
                Target := SafeCopy_1(945, y1, 3, 3, 1, true) ; target 확인
                Upper := SafeCopy_1(992, y1, 3, 3, 1, true) ; Upper 확인
                c := SafeCopy(1210, y1) ; sample수 확인
                Equipment := SafeCopy(1235, y1) ; 계측기 확인

                ; CMM이면 검사항목/세부내역 추가 스캔 (기준점은 함수 내에서 복원)
                isCMMRow := IsCMM(Equipment)
                if (isCMMRow)
                {
                    CharName := ScanCMMInfo(y1)
                    cmmValues := []
                }

                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Loop c
                {
                    Sleep(100)
                    result := Ave(Lower, Target, Upper, Equipment)
                    Send(result)
                    if (isCMMRow)
                        cmmValues.Push(result)
                    Sleep(100)
                    Send("{Tab}")
                }

                ; CMM 항목이면 수집 목록에 추가
                if (isCMMRow && cmmValues.Length > 0)
                {
                    fLower := 0, fTarget := 0, fUpper := 0
                    try fLower := Float(Lower)
                    try fTarget := Float(Target)
                    try fUpper := Float(Upper)
                    CMM_Records.Push({name: CharName, lower: fLower, target: fTarget, upper: fUpper, values: cmmValues})
                }

                ; 정량이고 샘플수가 5일 때만 좌우 스크롤
                if (c >= 5) {
                    Sleep(50)
                    SetDefaultMouseSpeed(15)
                    MouseClickDrag("Left", 1170, 809, 800, 809)
                    Sleep(300)  ; 웹 테이블 정렬 완료까지 대기
                    MouseClickDrag("Left", 1170, 809, 800, 809)
                    Sleep(300)  ; 웹 테이블 정렬 완료까지 대기
                    SetDefaultMouseSpeed(2)
                }
            }
            else if (DataValue = "정성" || DataValue = "첨부")
            {
                c := SafeCopy(1210, y) ; sample수 확인
                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Send("{Tab}")
                Sleep(50)
                Loop 1
                {
                    Sleep(100)
                    Send("OK")
                    Sleep(100)
                    Send("{Tab}")
                }
                ; 정성/첨부는 스크롤 불필요
            }
            else
            {
                MsgBox("오류. 화면을닫고 재시작하세요")
                return
            }

        }
    }
}
