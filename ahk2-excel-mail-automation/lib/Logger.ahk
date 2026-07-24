#Requires AutoHotkey v2.0
#Include Config.ahk
;==============================================================
; Logger.ahk
; 파일 로그를 남긴다. 매크로가 무인/장시간 실행되는 상황을
; 가정하므로, 오류 발생 시 원인을 추적할 수 있도록 모든 주요
; 단계와 예외를 기록한다.
;==============================================================

class Logger {
    static _ready := false

    static Init() {
        if Logger._ready
            return
        try DirCreate(Config.LogDir)
        Logger._ready := true
    }

    static Write(level, msg) {
        Logger.Init()
        line := FormatTime(, "yyyy-MM-dd HH:mm:ss") " [" level "] " msg "`n"
        try FileAppend(line, Config.LogFile, "UTF-8")
    }

    static Info(msg)  => Logger.Write("INFO", msg)
    static Warn(msg)  => Logger.Write("WARN", msg)
    static Error(msg) => Logger.Write("ERROR", msg)
}
