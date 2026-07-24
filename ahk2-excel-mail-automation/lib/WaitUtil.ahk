#Requires AutoHotkey v2.0
#Include Config.ahk
#Include Logger.ahk
;==============================================================
; WaitUtil.ahk
; "각 업무마다 작업이 완료되는 대기시간" 요구사항을 위한 공용
; 대기/재시도 헬퍼. 창 전환, 클립보드 준비, 일시적 오류에 대한
; 재시도를 한 곳에서 관리한다.
;==============================================================

class WaitUtil {
    ; 지정한 창(WinTitle 형식 문자열)이 나타날 때까지 기다렸다가 활성화한다.
    static ForWindow(criteria, timeoutSec := 0) {
        timeoutSec := timeoutSec ? timeoutSec : Config.WinWaitSec
        if !WinWait(criteria, , timeoutSec)
            throw TargetError("창을 찾지 못했습니다: " criteria)
        WinActivate(criteria)
        if !WinWaitActive(criteria, , timeoutSec)
            throw TargetError("창을 활성화하지 못했습니다: " criteria)
        Sleep(Config.WaitShort)
    }

    ; action(콜백, 인자 없음)을 실행하고 실패 시 설정된 횟수만큼 재시도한다.
    ; 모든 재시도가 실패하면 마지막 예외를 던진다.
    static Retry(action, desc := "", times := 0, delay := 0) {
        times := times ? times : Config.RetryCount
        delay := delay ? delay : Config.RetryDelay
        lastErr := ""
        loop times {
            try {
                return action()
            } catch as e {
                lastErr := e
                Logger.Warn(Format("{1} 재시도 {2}/{3}: {4}", desc, A_Index, times, e.Message))
                if A_Index < times
                    Sleep(delay)
            }
        }
        throw Error((desc ? desc " " : "") "실패(재시도 " times "회 초과): " (lastErr ? lastErr.Message : ""))
    }
}
