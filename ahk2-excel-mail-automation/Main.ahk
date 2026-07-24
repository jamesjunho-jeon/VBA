#Requires AutoHotkey v2.0
#SingleInstance Force
SetTitleMatchMode(2)
SetWorkingDir(A_ScriptDir)

#Include lib\Config.ahk
#Include lib\Logger.ahk
#Include lib\WaitUtil.ahk
#Include lib\ErrorHandler.ahk
#Include lib\InputCollector.ahk
#Include lib\ExcelController.ahk
#Include lib\MailController.ahk
#Include lib\Workflow.ahk
;==============================================================
; Main.ahk - 엑셀 소팅 -> 이미지 캡쳐 -> 메일 본문 붙여넣기 -> 발송
;
;   F1  : 매크로 시작 (2~17번 업무 프로세스 실행)
;   Esc : 매크로 중단 및 ahk2 스크립트 종료
;
; 실행 전 준비사항(1번 전제):
;   - 대상 통합문서를 열고, 제목+표로 구성된 시트를 활성화해 둔다.
;   - 표에는 자동 필터가 이미 설정되어 있어야 한다.
;   - 필터 기준값 목록이 있는 시트가 같은 통합문서 내에 있어야 한다.
;   - "메일" 이라는 제목(또는 포함 문자열)의 엣지 창이 열려 있어야 한다.
;==============================================================

global g_Running := false

F1:: StartAutomation()
Esc:: StopAndExit()

StartAutomation(*) {
    global g_Running
    if g_Running {
        ToolTip("이미 매크로가 실행 중입니다.")
        SetTimer(() => ToolTip(), -1500)
        return
    }

    g_Running := true
    try {
        Workflow.Run()
    } catch as e {
        Logger.Error("예기치 못한 오류: " e.Message " (" e.File ":" e.Line ")")
        MsgBox("예기치 못한 오류로 매크로가 중단되었습니다.`n" e.Message, "오류", "Icon!")
    } finally {
        g_Running := false
    }
}

StopAndExit(*) {
    Logger.Info("ESC 입력으로 매크로 중단 및 스크립트 종료")
    ; 보완: 실행 도중 Esc로 즉시 종료하는 경우, Workflow.Run()의 finally를
    ; 거치지 않고 프로세스가 바로 죽으므로 여기서 별도로 필터를 해제한다.
    if IsObject(Workflow.ActiveExcel) {
        try Workflow.ActiveExcel.ClearFilter(Workflow.ActiveSheet)
    }
    ExitApp()
}
