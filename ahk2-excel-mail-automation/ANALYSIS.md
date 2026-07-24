# 설계 분석 (ANALYSIS)

업무 정의: 엑셀에서 소팅/필터된 표를 이미지로 캡쳐하여 메일 본문에 순서대로
붙여넣고 발송하는 AHK2 자동화의 모듈화 구조와 설계 근거를 정리한다.

## 1. 확정된 두 가지 핵심 동작 (요구사항 원문이 두 가지로 해석 가능하여 사용자 확인을 거침)

| 항목 | 확정된 동작 |
|---|---|
| 9번(필터 걸기) | 필터시트 A열의 값을 **한 값씩 순차적으로** 필터 조건에 적용. 값 하나 = 라운드 하나. |
| 7번(반복횟수) | 같은 이미지를 여러 번 붙여넣는 게 아니라, **라운드마다 달라지는 캡쳐 이미지**를 본문에 붙여넣는 총 횟수. 즉 16번의 "필터 데이터 개수"와 사실상 동일한 개념. |

이에 따라 루프는 **1중 루프**(필터값 개수만큼 반복)로 구현했다. 사용자가 입력한
반복횟수(⑦)는 확인 절차의 참고값·불일치 경고용으로만 쓰고, 실제 루프 횟수는
필터시트에서 자동 계산한 실제 데이터 개수를 기준으로 한다(17번: "반복횟수가
안 되더라도 마지막 데이터 행이 되면 발송 후 종료"를 그대로 반영).

## 2. 요구사항 단계 → 모듈/함수 매핑

| 단계 | 업무 정의 | 담당 모듈 | 함수 |
|---|---|---|---|
| 1 | 엑셀 초기 구성(전제) | ExcelController | `Attach()` (AutoFilter 존재 검증) |
| 2~7 | 각 입력값 수집 | InputCollector | `_PromptForm()` |
| 8 | 입력사항 확인/재입력 | InputCollector | `_Confirm()`, `Collect()` 루프 |
| 9 | 캡쳐자동화(필터 걸기) | ExcelController | `ApplyFilter()` |
| 10, 14 | 영역 복사(반복) | ExcelController | `GetCopyRange()`, `CopyRangeAsPicture()` |
| 11 | 메일발송 열기 | MailController | `ActivateMailApp()`, `OpenNewMail()` |
| 12 | 메일제목 출력 | MailController | `TypeSubject()` |
| 13 | 메일주소 출력 | MailController | `TypeAddress()` |
| 15 | 붙여넣기(Ctrl+V, ↓, Enter) | MailController | `PasteAndAdvance()` |
| 16 | 9~15 반복, 완료 메시지 | Workflow | `Run()` 메인 루프 |
| 17 | 최종 발송 | MailController / Workflow | `SendMail()` |
| 전체 | F1 시작 / Esc 종료 | Main.ahk | 핫키 바인딩 |
| 보완 | 대기시간, 재시도, 오류처리 | WaitUtil, ErrorHandler, Logger | 전 모듈 공용 |
| 보완 | 취소/완료 시 필터 해제(전체 목록 표시) | ExcelController, Workflow, Main.ahk | `ClearFilter()`, `Workflow.Run()`의 finally, `StopAndExit()` |

## 3. 모듈 구조 (라이브러리화)

```
ahk2-excel-mail-automation/
  Main.ahk              엔트리포인트: 핫키(F1/Esc)만 정의, 실제 로직은 lib 참조
  lib/
    Config.ahk           튜닝 가능한 값(대기시간, 창 제목, 단축키) - 유일한 설정 지점
    Logger.ahk            파일 로그 (logs/mail_YYYYMMDD.log)
    WaitUtil.ahk          창 대기, 재시도(Retry) 공용 유틸
    ErrorHandler.ahk      오류 발생 시 사용자 대응(건너뛰기/중단) 정책
    InputCollector.ahk    2~8번: 입력 GUI + 검증 + 확인
    ExcelController.ahk   1/9/10/14번: Excel COM 제어
    MailController.ahk    11/12/13/15/17번: 메일 창 키보드 제어
    Workflow.ahk          전체 오케스트레이션(9~17번)
```

**설계 원칙**
- `Config.ahk` 외 다른 파일에는 매직넘버(대기시간, 창 제목 등)를 두지 않는다.
  회사 PC 환경이나 메일 서비스가 바뀌어도 이 파일만 고치면 된다.
- `ExcelController` / `MailController`는 서로를 모른다(단방향 의존 없음).
  Excel 자동화만 재사용하고 싶으면 `ExcelController`만 `#Include`하면 된다.
- `Workflow`만 두 컨트롤러를 조합한다 → 업무 순서가 바뀌면 `Workflow.ahk`만 수정.
- 다른 업무(예: 다른 표, 다른 메일 서비스)에 재사용하려면 `InputCollector`와
  `Config`만 교체하면 되도록 분리했다.

## 4. 오류 처리 정책

| 구간 | 처리 방식 |
|---|---|
| 입력 취소(Gui 닫기/취소) | 매크로 조용히 종료 (재실행은 F1으로) |
| Excel 미실행/시트 없음/자동필터 없음 | `ErrorHandler.Fatal()` - 즉시 안내 후 중단 |
| 필터시트 없음/데이터 0건 | `ErrorHandler.Fatal()` |
| 필터 결과 0건(빈 셀) | 오류 아님. 해당 라운드만 `continue`(14번 "빈 셀은 건너뛰어") |
| 라운드 중 일시적 오류(창 전환 실패, 클립보드 지연 등) | `WaitUtil.Retry()`로 자동 재시도(기본 3회, 1초 간격) |
| 재시도 모두 실패 | `ErrorHandler.AskSkipOrAbort()` - 사용자에게 건너뛰기/중단 확인 |
| 전 라운드 실패(발송할 이미지 0건) | 메일 발송 생략 + 오류 안내 |
| 발송 단계 실패 | `ErrorHandler.Fatal()` |
| 스크립트 전역 미처리 예외 | `Main.ahk`의 try/catch가 최종 방어선 |
| Esc 입력 | 즉시 `ExitApp()` - 진행 중 COM 작업/메일 임시본이 남을 수 있음(아래 5번 참고) |

## 5. 대기시간 설계 (Config.ahk)

| 상수 | 값 | 용도 |
|---|---|---|
| WaitShort | 300ms | 탭 이동, 키 입력 사이 |
| WaitMedium | 800ms | 시트 전환, 필터 재계산, 창 전환 |
| WaitLong | 1500ms | 새 메일 창 로딩, 이미지 붙여넣기 렌더링 |
| WinWaitSec | 10초 | 창이 뜰 때까지 최대 대기 |
| ClipWaitSec | 5초 | 클립보드(그림) 준비 대기 |
| RetryCount / RetryDelay | 3회 / 1초 | 라운드 실패 시 재시도 |

값은 실제 PC 사양·네트워크·엑셀 데이터 크기에 따라 조정이 필요할 수 있으므로
전부 `Config.ahk` 한 곳에 모아뒀다.

## 6. 알려진 가정 및 제약 (실행 전 확인 필요)

1. **"메일" 창 식별**: `Config.EdgeMailTitle := "메일"`은 부분일치(SetTitleMatchMode 2)
   로 엣지 창 제목에 "메일"이 포함되는지 확인한다. 실제 서비스(Outlook Web,
   Gmail 등) 창 제목이 다르면 이 값을 수정해야 한다.
2. **새 메일 단축키**: Shift+N은 Outlook Web 등에서 흔히 쓰이는 단축키다.
   사용 중인 메일 서비스가 다르면 `Config.NewMailHotkey`를 바꿔야 한다.
3. **발송 단축키**: Ctrl+Enter로 가정했다(Outlook Web/Gmail 공통). 다르면
   `Config.SendMailHotkey` 수정.
4. **주소 입력 후 Tab 3회**: 실제 메일 UI의 수신자/참조/제목 탭 순서에 따라
   `Config.TabsAfterAddress` 값 조정이 필요할 수 있다.
5. **여러 줄 헤더**: Excel의 AutoFilter는 구조적으로 헤더를 1행만 인식한다.
   표 위에 제목/설명이 여러 줄 있어도 시작셀(②)만 정확히 지정하면 캡쳐 범위
   계산에는 영향이 없다. (즉 "여러줄 헤더"는 시작셀보다 위쪽 영역으로 처리됨)
6. **필터 조건 타입**: `AutoFilter`에 문자열 형태로 값을 전달한다. 날짜/서식이
   특수한 값은 매치가 안 될 수 있어 별도 포맷 변환이 필요할 수 있다.
7. **Esc 즉시 종료**: 요구사항대로 확인 없이 즉시 스크립트를 종료한다.
   `Workflow.ActiveExcel/ActiveSheet`를 통해 필터 해제는 시도하지만(6-8번
   참고), 작성 중이던 메일 임시 저장본은 그대로 남을 수 있다. 또한 AHK 특성상
   Excel COM 호출이 실제로 실행되는 그 찰나(순수 Send/Sleep 구간이 아닌)에
   Esc가 눌리면 정리 코드가 끼어들 타이밍을 못 잡을 수 있어 100% 보장은 아니다
   (최선 노력 방식). 안전한 되돌리기가 필요하면 종료 전 확인창을 추가하는
   것을 권장한다(현재는 미적용).
8. **필터 해제(보완 기능)**: 작업이 취소되거나(입력 취소 이후 단계에서 오류로
   중단, 라운드 중 "전체중단" 선택, Esc 즉시 종료 포함) 정상 완료되면
   `ExcelController.ClearFilter()`가 `Worksheet.ShowAllData()`를 호출해
   필터 조건만 해제하고(자동필터 드롭다운 자체는 유지) 전체 목록을 다시
   표시한다. 8번(입력 확인 단계)에서 "취소"한 경우는 아직 엑셀 필터를
   건드리기 전이므로 별도 처리가 필요 없다.

## 7. 확장 포인트

- 다른 표/다른 메일 서비스에 재사용: `Config.ahk`만 교체.
- 필터 기준이 "값 목록"이 아니라 "범위 조건(예: 이상/이하)"으로 바뀌면
  `ExcelController.ApplyFilter()`의 `AutoFilter(fieldIndex, value)` 호출부만 수정.
- 여러 명에게 각각 다른 라운드를 따로 발송(현재는 1건의 메일에 여러 이미지를
  이어붙이는 구조)하려면 `Workflow.Run()`의 `mailStarted` 분기를 라운드마다
  새 메일을 여는 구조로 바꾸면 된다.
- GUI 로그 대신 진행률 표시가 필요하면 `Logger`에 `ToolTip`/`Progress` Gui를
  추가하는 방식으로 확장 가능(현재는 로그 파일 + 완료 MsgBox만 제공).
