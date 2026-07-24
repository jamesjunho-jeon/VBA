#Requires AutoHotkey v2.0
;==============================================================
; Config.ahk
; 매크로 전역 설정값 - 대기시간, 창 식별자, 단축키 등을
; 한 곳에서 관리한다. 환경(회사 PC, 메일 서비스 등)이 바뀌면
; 이 파일만 수정하면 되도록 모든 "튜닝 가능한 값"을 모았다.
;==============================================================

class Config {
    ; ---- 매크로 제어 단축키 ----
    static StartHotkey := "F1"
    static StopHotkey  := "Esc"

    ; ---- 대기시간(ms) : 각 업무 단계가 안정화될 때까지의 대기 ----
    static WaitShort   := 300     ; 짧은 UI 반응 대기 (탭 이동, 키 입력 사이)
    static WaitMedium  := 800     ; 창 전환, 필터 적용 등 중간 대기
    static WaitLong    := 1500    ; 새 메일 로딩, 붙여넣기 렌더링 등 긴 대기

    static WinWaitSec  := 10      ; 창 활성화 대기 제한(초)
    static ClipWaitSec := 5       ; 클립보드 준비 대기 제한(초)

    static RetryCount  := 3       ; 라운드(필터값 1건) 처리 실패 시 공통 재시도 횟수
    static RetryDelay  := 1000    ; 재시도 간 대기(ms)

    ; ---- 대상 애플리케이션 식별자 ----
    static ExcelProcess   := "ahk_exe EXCEL.EXE"

    ; 11번(메일발송 열기): "메일" 이라는 헤더를 가진 엣지 창
    ; 실제 창 제목이 다르면(예: "메일 - Outlook", "받은편지함 - 메일") 이 값만 수정한다.
    static EdgeMailTitle  := "메일"
    static EdgeProcess    := "ahk_exe msedge.exe"

    ; ---- 메일 편집기 동작 키 ----
    static NewMailHotkey  := "+n"        ; Shift+N : 새 메일 작성 (11번)
    static SendMailHotkey := "^Enter"    ; Ctrl+Enter : 메일 발송 (17번, Outlook Web/Gmail 공통)
    static TabsAfterAddress := 3         ; 13번: 주소 입력 후 본문으로 이동하기 위한 Tab 횟수

    ; ---- 로그 파일 ----
    static LogDir  := A_ScriptDir "\logs"
    static LogFile => Config.LogDir "\mail_" FormatTime(, "yyyyMMdd") ".log"
}
