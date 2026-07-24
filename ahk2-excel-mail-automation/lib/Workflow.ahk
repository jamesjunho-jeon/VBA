#Requires AutoHotkey v2.0
#Include Config.ahk
#Include Logger.ahk
#Include WaitUtil.ahk
#Include ErrorHandler.ahk
#Include InputCollector.ahk
#Include ExcelController.ahk
#Include MailController.ahk
;==============================================================
; Workflow.ahk
; 업무 정의: 전체 프로세스(1~17번)의 오케스트레이션.
;
; 확정된 동작 방식(사용자 확인 완료):
;   - 9번(필터): 필터시트 A열의 값을 "한 개씩 순차적으로" 필터
;     조건으로 걸며, 값마다 별도의 캡쳐→붙여넣기 라운드를 만든다.
;   - 7번(반복횟수): "본문에 복사해 붙여넣는 횟수"로, 라운드마다
;     서로 다른(그 라운드의 필터 결과) 이미지를 붙여넣는다.
;   - 실제 루프 횟수는 필터시트의 실제 데이터 개수(16번)를
;     기준으로 하고, 사용자가 입력한 반복횟수(⑦)는 확인/경고용
;     참고값으로 사용한다. 실제 데이터가 먼저 끝나면 그 시점에
;     발송하고 종료한다(17번).
;   - 보완: 작업이 취소되거나(오류/사용자중단 포함) 완료되면
;     필터 조건을 해제하여 원래의 전체 목록이 다시 보이도록 한다.
;     (Esc로 즉시 종료하는 경로는 Main.ahk 참고 - ActiveExcel/
;      ActiveSheet를 통해 별도로 정리한다)
;==============================================================

class Workflow {
    ; Esc 등 외부에서 즉시 종료할 때도 필터를 해제할 수 있도록
    ; 현재 실행 중인 엑셀 컨트롤러/시트를 노출해 둔다.
    static ActiveExcel := ""
    static ActiveSheet := ""

    static Run() {
        Logger.Info("=== 자동화 시작 ===")

        ; ---- 2~8번: 입력 수집 및 확인 ----
        try {
            userInput := InputCollector.Collect()
        } catch as e {
            Logger.Warn("입력 취소: " e.Message)
            return
        }

        ; ---- 1번 전제 확인 + 3번 필터 목록 로드 ----
        excel := ExcelController()
        try {
            mainSheet := excel.Attach()
            filterList := excel.GetFilterList(userInput["FilterSheet"])
        } catch as e {
            ErrorHandler.Fatal("엑셀 준비 중 오류: " e.Message)
            return
        }

        ; 이 시점부터 필터를 건드리므로, 어떤 경로로 끝나든(완료/오류/취소)
        ; finally 에서 반드시 필터를 해제한다.
        Workflow.ActiveExcel := excel
        Workflow.ActiveSheet := mainSheet
        try {
            Workflow._Execute(excel, mainSheet, filterList, userInput)
        } finally {
            excel.ClearFilter(mainSheet)   ; 보완: 취소/완료 시 전체 목록 복원
            Workflow.ActiveExcel := ""
            Workflow.ActiveSheet := ""
        }
    }

    static _Execute(excel, mainSheet, filterList, userInput) {
        actualCount := filterList.Length
        declaredCount := userInput["RepeatCount"]

        if actualCount = 0 {
            ErrorHandler.Fatal("필터시트(" userInput["FilterSheet"] ")에 필터 데이터가 없습니다.")
            return
        }
        if actualCount != declaredCount
            Logger.Warn(Format("입력한 반복횟수({1})와 필터시트의 실제 데이터 개수({2})가 다릅니다. 실제 데이터 개수를 기준으로 진행합니다.", declaredCount, actualCount))

        loopCount := actualCount
        mail := MailController()
        mailStarted := false
        sentAny := false

        ; 라운드 1건(필터값 1개) 처리: 9번(필터) -> 10/14번(복사) -> 11~13/15번(메일)
        RoundBody(idx, value) {
            excel.ActivateSheet(mainSheet)
            matched := excel.ApplyFilter(mainSheet, userInput["FilterCell"], value)
            if !matched
                return false   ; 14번: 빈 셀(매치 없음)은 건너뛴다

            copyRange := excel.GetCopyRange(mainSheet, userInput["StartCell"])
            excel.CopyRangeAsPicture(copyRange)

            if !mailStarted {
                mail.ActivateMailApp()          ; 11번
                mail.OpenNewMail()               ; 11번
                mail.TypeSubject(userInput["Subject"])       ; 12번
                mail.TypeAddress(userInput["MailAddress"])   ; 13번
                mailStarted := true
            } else {
                mail.ActivateMailApp()
            }

            mail.PasteAndAdvance()               ; 15번
            return true
        }

        ; ---- 16번: 필터 데이터 개수만큼 9~15번 반복 ----
        loop loopCount {
            idx := A_Index
            value := filterList[idx]
            try {
                matched := WaitUtil.Retry(() => RoundBody(idx, value), "라운드 " idx " (필터값: " value ")")
                if matched {
                    sentAny := true
                    Logger.Info(Format("[{1}/{2}] 필터값 '{3}' 처리 완료", idx, loopCount, value))
                } else {
                    Logger.Warn(Format("[{1}/{2}] 필터값 '{3}' 에 해당하는 데이터가 없어 건너뜁니다.", idx, loopCount, value))
                }
            } catch as e {
                if ErrorHandler.AskSkipOrAbort(idx, value, e) = "abort" {
                    Logger.Warn("사용자 요청으로 자동화를 중단합니다.")
                    break
                }
            }
        }

        if !sentAny {
            ErrorHandler.Fatal("본문에 붙여넣을 데이터가 없어 메일을 발송하지 않았습니다.")
            return
        }

        ; ---- 17번: 마지막 라운드가 끝나면(반복횟수 도달 여부와 무관) 메일 발송 ----
        try {
            mail.ActivateMailApp()
            mail.SendMail()
        } catch as e {
            ErrorHandler.Fatal("메일 발송 중 오류: " e.Message)
            return
        }

        Logger.Info("=== 자동화 완료 ===")
        MsgBox("메일 발송이 완료되었습니다. (총 " loopCount "건 중 처리)", "완료", "Iconi")
    }
}
