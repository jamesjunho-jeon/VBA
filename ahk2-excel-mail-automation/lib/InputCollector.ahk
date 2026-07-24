#Requires AutoHotkey v2.0
#Include Config.ahk
;==============================================================
; InputCollector.ahk
; 업무 정의: 2~8번 단계를 담당한다.
;   2. 시작셀 입력   3. 필터시트 입력   4. 필터셀주소 입력
;   5. 메일제목 입력  6. 메일주소 입력   7. 반복횟수 입력
;   8. 입력사항 확인 (Yes 면 진행, No 면 값을 보존한 채 재입력)
; 모든 값은 형식 검증을 통과해야 확인 단계로 넘어간다.
;==============================================================

class InputCollector {
    ; 최종 확정된 입력값을 Map 으로 반환한다. 사용자가 취소하면 예외를 던진다.
    static Collect() {
        prev := ""
        loop {
            data := InputCollector._PromptForm(prev)
            if !data
                throw Error("사용자가 입력을 취소했습니다.")
            if InputCollector._Confirm(data)
                return data
            prev := data   ; 8번: "아니오" -> 이전 입력값을 유지한 채 다시 입력받는다
        }
    }

    static _PromptForm(prev) {
        data := ""
        done := false
        ok := false

        g := Gui("+AlwaysOnTop", "① 작업 정보 입력")
        g.SetFont("s10")

        g.AddText("w320", "② 시작셀  (캡처 범위가 시작되는 제목 셀, 예: B3)")
        eStart := g.AddEdit("w300", prev ? prev["StartCell"] : "")

        g.AddText("w320", "③ 필터시트 이름  (필터 기준값 목록이 있는 시트)")
        eFSheet := g.AddEdit("w300", prev ? prev["FilterSheet"] : "")

        g.AddText("w320", "④ 필터셀 주소  (필터를 걸 표의 헤더 셀, 예: C5)")
        eFCell := g.AddEdit("w300", prev ? prev["FilterCell"] : "")

        g.AddText("w320", "⑤ 메일 제목")
        eSubject := g.AddEdit("w300", prev ? prev["Subject"] : "")

        g.AddText("w320", "⑥ 메일 주소  (여러 명은 , 또는 ; 로 구분)")
        eMail := g.AddEdit("w300", prev ? prev["MailAddress"] : "")

        g.AddText("w320", "⑦ 반복횟수  (본문에 붙여넣기할 횟수)")
        eRepeat := g.AddEdit("w300", prev ? String(prev["RepeatCount"]) : "")

        errText := g.AddText("w320 cRed", "")

        btnOk := g.AddButton("w140 Default", "확인")
        btnCancel := g.AddButton("w140 x+10", "취소")

        Submit(*) {
            v := InputCollector._Validate(eStart.Text, eFSheet.Text, eFCell.Text,
                eSubject.Text, eMail.Text, eRepeat.Text)
            if v.ok {
                data := v.data
                ok := true
                done := true
            } else {
                errText.Text := v.err
            }
        }
        Cancel(*) {
            done := true
        }

        btnOk.OnEvent("Click", Submit)
        btnCancel.OnEvent("Click", Cancel)
        g.OnEvent("Close", Cancel)
        g.OnEvent("Escape", Cancel)

        g.Show()
        while !done
            Sleep(50)
        g.Destroy()

        return ok ? data : false
    }

    static _Validate(startCell, fSheet, fCell, subject, mail, repeatStr) {
        startCell := Trim(startCell)
        fSheet := Trim(fSheet)
        fCell := Trim(fCell)
        subject := Trim(subject)
        mail := Trim(mail)
        repeatStr := Trim(repeatStr)

        cellPattern := "^\$?[A-Za-z]{1,3}\$?\d{1,7}$"

        if !RegExMatch(startCell, cellPattern)
            return {ok: false, err: "시작셀 주소 형식이 올바르지 않습니다. 예) B3"}
        if fSheet = ""
            return {ok: false, err: "필터시트 이름을 입력하세요."}
        if !RegExMatch(fCell, cellPattern)
            return {ok: false, err: "필터셀 주소 형식이 올바르지 않습니다. 예) C5"}
        if subject = ""
            return {ok: false, err: "메일 제목을 입력하세요."}
        if !InputCollector._ValidateEmails(mail)
            return {ok: false, err: "메일 주소 형식이 올바르지 않습니다."}
        if !RegExMatch(repeatStr, "^\d+$") || Integer(repeatStr) < 1
            return {ok: false, err: "반복횟수는 1 이상의 숫자여야 합니다."}

        data := Map(
            "StartCell", StrUpper(startCell),
            "FilterSheet", fSheet,
            "FilterCell", StrUpper(fCell),
            "Subject", subject,
            "MailAddress", mail,
            "RepeatCount", Integer(repeatStr)
        )
        return {ok: true, data: data}
    }

    static _ValidateEmails(text) {
        if text = ""
            return false
        for part in StrSplit(text, [";", ","]) {
            part := Trim(part)
            if part = ""
                continue
            if !RegExMatch(part, "^[^@\s]+@[^@\s]+\.[^@\s]+$")
                return false
        }
        return true
    }

    ; 8번: 입력사항 확인
    static _Confirm(data) {
        msg := "다음 내용으로 진행할까요?`n`n"
            . "② 시작셀: " data["StartCell"] "`n"
            . "③ 필터시트: " data["FilterSheet"] "`n"
            . "④ 필터셀: " data["FilterCell"] "`n"
            . "⑤ 메일 제목: " data["Subject"] "`n"
            . "⑥ 메일 주소: " data["MailAddress"] "`n"
            . "⑦ 반복횟수: " data["RepeatCount"]

        result := MsgBox(msg, "입력 확인", "YesNo Icon?")
        return (result = "Yes")
    }
}
